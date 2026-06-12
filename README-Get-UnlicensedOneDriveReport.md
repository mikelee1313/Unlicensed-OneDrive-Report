# Get-UnlicensedOneDriveReport.ps1

PowerShell script to identify and report unlicensed OneDrive accounts across a Microsoft 365 tenant (including multi-geo), with optional archive cost projection and optional email alerts.

This script primarily uses Microsoft Graph application permissions and can also download the SharePoint Admin unlicensed OneDrive report (ExportToCSV) for archive-focused enrichment.

## UnlicensedOneDrive Report Script Summary

This PowerShell script identifies and reports on **unlicensed OneDrive accounts** across a Microsoft 365 tenant, including multi-geo environments. It's designed to help organizations comply with Microsoft's enforcement timeline for unlicensed OneDrive accounts (effective January 27, 2025) and track the **post-archive deletion timeline** introduced in **Change Notice MC1381110**.

### What It Does

The script discovers three populations of unlicensed OneDrive accounts:

1. **Active Entra ID Users** — Users with active accounts but no enabled OneDrive/SharePoint license
2. **Soft-Deleted Users** — Users in the Entra ID 30-day recycle bin (whose OneDrives still exist)
3. **Archived OneDrive Sites** — Personal OneDrive sites already archived by Microsoft after the owner was purged from Entra ID (>30 days post-deletion)

For each account, it collects:
- User/site identity and display names
- License status and removal dates (via audit logs)
- OneDrive storage usage (GB) and URL
- Calculated timelines for read-only, archive, and deletion phases

### Critical Deletion Risk Timeline (MC1381110)

The script enforces Microsoft's documented archival timeline and **deletion risk window** as defined in **Change Notice MC1381110**:

| Phase | Days from Unlicensed | Status | Risk |
|-------|---|---|---|
| **Read-Only** | Day 60 | OneDrive becomes read-only | Sites are inaccessible for writes |
| **Archived** | Day 93 | OneDrive archived by Microsoft | Minimal metadata queries; high retrieval cost |
| **Deletion Window** | Days 94–365+ | ⚠️ **DELETION RISK (MC1381110)** | If PAYG not enabled, archived accounts are deleted after ~365 days |

#### Deletion Risk Details (MC1381110)
- **If PAYG is NOT enabled** (default): Archived unlicensed OneDrive accounts are subject to **automatic deletion** after approximately **365 days** from the unlicensed date per MC1381110, even if a retention hold exists.
- **If PAYG is enabled**: Archived accounts persist indefinitely but incur **$0.05/GB/month** storage fees and **$0.60/GB** one-time reactivation costs.
- **Manual Config**: The script uses a manual setting (`$PayGEnabledForUnlicensedOneDrive`) since Microsoft has no API to read PAYG status.

The script calculates **DaysUntilDeletion** for each archived account (configurable via `$ArchiveDeletionThresholdDays`, defaulting to 365) and sends alert emails when deletion risk approaches the configured threshold (default: 30 days before Day 365).

## What This Script Does

The script reports three account populations:

1. Active Entra ID users without an enabled OneDrive/SharePoint service plan.
2. Soft-deleted Entra ID users (within recycle bin) whose OneDrive may still exist.
3. Archived unlicensed OneDrive accounts discovered from SharePoint Admin report download (or Graph Sites mode if explicitly selected).

For each discovered account/site, it can include:

- Identity fields (UPN, display name, account/source type)
- Unlicensed reason and date (when available)
- Timeline milestones (Day 60 read-only, Day 93 archive)
- Deletion risk timeline based on your manual PAYG setting
- Storage metrics and cost projections
- Deletion blocked reason (from downloaded SPO report)

## Key Capabilities

- **Multi-geo support** — Single app registration and token covers all datacenters (NAM, APC, CAN, DEU, GBR, IND, JPN, etc.)
- **MC1381110 compliance reporting** — Tracks post-archive deletion timeline and PAYG reactivation risk
- **Audit log integration** — Bulk query for license-change dates (optional, requires `AuditLog.Read.All`)
- **SharePoint Admin integration** — Downloads existing unlicensed reports from SPO admin portals
- **Email alerts** — Notifies admins of approaching read-only, archive, and deletion milestones
- **Cost estimation** — Calculates storage fees and reactivation costs based on tenant PAYG status
- **Throttle handling** — Built-in retry logic with exponential backoff for Graph API rate limits
- **Graph API only** — No SPO PowerShell module or per-geo token management required
- Bulk user enumeration with paging
- Graph batch API for faster drive lookups
- Optional merge of downloaded SPO report rows into final dataset
- Optional HTML email notifications for upcoming read-only/archive windows

## Important Design Decisions

- PAYG status is manual (`$PayGEnabledForUnlicensedOneDrive`).
  - There is no reliable public API/cmdlet readback for the unlicensed OneDrive billing toggle.
- In `SPODownload` mode, there is no automatic fallback to Graph Sites when download returns no rows.
  - This avoids expensive per-site Graph checks and keeps runtime predictable.

## Prerequisites

### Runtime

- PowerShell 7+ recommended (Windows supported).
- Network access to:
  - `https://graph.microsoft.com`
  - `https://login.microsoftonline.com`
  - SharePoint admin endpoints in `$SPOAdminUrls`

### App Registration and Permissions

Use an Entra app registration with application permissions:

**Required for core report:**

- `User.Read.All`
- `Directory.Read.All`
- `Files.Read.All`

**Optional features:**

- `AuditLog.Read.All` (license-removal date enrichment)
- `Sites.Read.All` (if `GraphSites` mode is used)
- `Mail.Send` (email notifications)

Admin consent must be granted for the configured permissions.

### Authentication

The script supports:

- Certificate auth (`$AuthType = 'Certificate'`) - recommended
- Client secret auth (`$AuthType = 'ClientSecret'`)

If using certificate auth, ensure the cert exists at:

- `Cert:\LocalMachine\My\<thumbprint>` or `Cert:\CurrentUser\My\<thumbprint>`

## Configuration

All configuration is in the script under `CONFIGURATION SECTION`.

Most important settings:

- Tenant/app/auth:
  - `$tenantId`
  - `$clientId`
  - `$AuthType`
  - `$Thumbprint` or `$clientSecret`
- Archive timelines:
  - `$ReadOnlyThresholdDays` (default 60)
  - `$ArchiveThresholdDays` (default 93)
  - `$ArchiveDeletionThresholdDays` (default 365)
- PAYG (manual):
  - `$PayGEnabledForUnlicensedOneDrive`
- Archived collection mode:
  - `$GetCurrentlyArchived`
  - `$ArchivedCollectionMode` (`SPODownload` or `GraphSites`)
  - `$SPOAdminUrls`
- Data enrichment toggles:
  - `$includeLicenseRemovalDates`
  - `$IncludeDownloadedRowsInMainReport`
- Notifications:
  - `$SendEmailNotifications`
  - `$EmailTo`
  - `$EmailFrom`
  - `$DaysToNotifyBeforeReadOnly`
  - `$DaysToNotifyBeforeArchive`
  - `$DaysToNotifyBeforeDeletion`

## Mode Guidance

### SPODownload (recommended default)

- Uses SharePoint Admin ExportToCSV.
- Best for tenant-scale reporting speed and includes fields like deletion-block reasons.
- No automatic fallback to Graph Sites when no rows are returned.

### GraphSites

- Enumerates personal sites via Graph Sites APIs.
- Useful if SPO export path is unavailable.
- Can be slower and more expensive at large scale.

## How to Run

From PowerShell in the script directory:

```powershell
.\Get-UnlicensedOneDriveReport.ps1
```

The output CSV is written to `$OutputFolder` (defaults to `%TEMP%`) as:

- `UnlicensedOneDrive_<timestamp>.csv`

If SPO report merging is enabled, an intermediate merged file may be created and then cleaned up by the script.

## Output Schema (Main CSV)

Typical output columns include:

- `UserSource` (`Active`, `SoftDeleted`, `Archived`)
- `DisplayName`
- `UserPrincipalName`
- `AccountEnabled`
- `UnlicensedDueTo`
- `UnlicensedDate`
- `DaysSinceUnlicensed`
- `ReadOnlyDate`
- `ArchiveDate`
- `DaysUntilReadOnly`
- `DaysUntilArchive`
- `DeletionBlockedBy` (from SPO downloaded report rows when available)
- `DaysUntilDeletion`
- `UrgencyStatus`
- `StorageUsedGB`
- `StorageTotalGB`
- `ProjMonthlyStorageCost`
- `ProjReactivationCost`
- `DriveUrl`
- `DriveLastModified`
- `Notes`

Notes:

- `DeletionBlockedBy` is typically populated for accounts sourced or backfilled from SPO download data.
- For archived accounts without known unlicensed date, date-driven counters may remain unknown or use archive-specific status labels.

## Sample Output:

<img width="1608" height="158" alt="image" src="https://github.com/user-attachments/assets/67d289fd-c2e9-40f4-86af-1c39dc48ac38" />

<img width="1542" height="161" alt="image" src="https://github.com/user-attachments/assets/44b7ee81-a9a3-4c42-8741-1af09a488dd1" />

<img width="1865" height="161" alt="image" src="https://github.com/user-attachments/assets/3308163e-707e-4bf9-a374-9cc30fec990e" />


## E-Mail Notification Samples:

<img width="1249" height="261" alt="image" src="https://github.com/user-attachments/assets/f25bf2e0-4f98-49e2-8cc5-dd11f750276c" />

<img width="1248" height="485" alt="image" src="https://github.com/user-attachments/assets/8ef6b545-7fd5-4e24-a7fe-40d55294aed3" />

<img width="1252" height="474" alt="image" src="https://github.com/user-attachments/assets/26dfaa54-f460-42a7-8b24-bb2178b3f160" />


## Processing Flow

1. Acquire Graph token.
2. Enumerate active unlicensed users.
3. Enumerate soft-deleted users.
4. Collect archived accounts (`SPODownload` or `GraphSites`).
5. Optionally enrich active users with audit log license-removal dates.
6. Resolve drive metadata (Graph batch + fallback sequential slice retry).
7. Merge/backfill rows from downloaded SPO reports.
8. Apply milestone, urgency, deletion-risk, and cost enrichment.
9. Export final CSV.
10. Optionally send email notifications.

## Performance Notes

- Graph drive lookups use JSON batching (up to 20 requests per batch).
- Retry logic handles Graph throttling and transient errors.
- `SPODownload` mode avoids expensive per-site Graph enumeration.
- For very large tenants:
  - Keep `$delayBetweenRequests = 0` unless throttling patterns require it.
  - Consider reducing optional enrichments if runtime is a concern.

## Security Notes

- Prefer certificate auth over client secret.
- Limit app permissions to only what is required.
- Protect certificate private keys and rotate credentials regularly.
- Do not commit tenant IDs, client IDs, thumbprints, or email targets to public repositories unless intentionally sanitized.

## Troubleshooting

### Authentication failures

- Verify `tenantId`, `clientId`, and cert/secret values.
- Confirm cert exists in the configured store and has private key access.
- Verify app permissions and admin consent.

### Empty or low archived results in SPODownload mode

- Verify `$SPOAdminUrls` are correct and reachable.
- Confirm app has rights for SPO export path.
- Remember: SPODownload mode does not auto-fallback to Graph when no rows return.

### Missing UnlicensedDate for active users

- Ensure `AuditLog.Read.All` is granted.
- Increase `$AuditLogLookbackDays` as needed (max 180 in current script design).

### Email send failures

- Confirm `Mail.Send` application permission and admin consent.
- Ensure `$EmailFrom` is a licensed Exchange Online mailbox.

## Known Limitations

- Users deleted and purged long ago may be unrecoverable by APIs.
- Reconstructed UPN from personal site URL is best-effort (underscore/dot ambiguity).
- Some archive/deletion timing fields are inferred from policy thresholds and available dates.
- PAYG status is manual input, not dynamically discovered.

## Suggested Repository Structure

If publishing this script to GitHub, a clean structure could be:

- `Get-UnlicensedOneDriveReport.ps1`
- `README.md` (this document)
- `LICENSE`
- `docs/` (optional architecture notes/screenshots)
- `samples/` (optional sample CSV output with sanitized data)

## Disclaimer

This script is provided as-is. Validate in a non-production tenant first and review output carefully before operational decisions.

## Authors

- Mike Lee
- Mariel Williams
