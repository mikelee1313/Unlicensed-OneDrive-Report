# Get-UnlicensedOneDriveReport.ps1

A PowerShell script that identifies all unlicensed OneDrive accounts across a Microsoft 365 tenant using the **Microsoft Graph API only** — no SharePoint Online PowerShell module, no per-geo tokens, and no manual admin center navigation required.

The Get‑UnlicensedOneDriveReport.ps1 script scans the entire Microsoft 365 tenant and identifies all unlicensed OneDrive accounts using Microsoft Graph only—no SharePoint Online PowerShell module, no per‑geo tokens, and no manual admin center checks required. It detects unlicensed OneDrives across active users, soft‑deleted users, and already‑archived OneDrive sites, and calculates where each account falls in Microsoft’s Day 60 (read‑only) and Day 93 (archive/delete) enforcement timeline. The output is a prioritized CSV report with clear urgency labels

- Proactive risk management: Admins can quickly identify OneDrive accounts that are approaching read‑only or archive status, reducing the risk of data loss or surprise support escalations.
- Cost control: With Microsoft enforcing archiving and billing for unlicensed OneDrives, the report helps admins find accounts that may soon incur archive or storage charges and act before costs are triggered.
- Tenant‑wide visibility: The script is multi‑geo aware and reports across all regions through a single Graph token, something that’s otherwise difficult to do at scale using built‑in tools. 
- Automation‑ready: Designed for scheduled or repeat use, it enables SPO admins to incorporate unlicensed OneDrive monitoring into regular governance and offboarding processes

---

## Overview

When a Microsoft 365 user loses their OneDrive/SharePoint license (license removed, user deleted, or license plan disabled), Microsoft begins a countdown toward archival and eventual deletion of their OneDrive:

| Day | Action |
|-----|--------|
| Day 0 | License removed / user deleted |
| Day 60 | OneDrive goes **read-only** |
| Day 93 | OneDrive is **archived** (or deleted if billing is not enabled) |

> Microsoft began enforcing this policy on **January 27, 2025**.  
> Reference: [Unlicensed OneDrive accounts](https://learn.microsoft.com/en-us/sharepoint/unlicensed-onedrive-accounts)

This script scans the entire tenant, identifies **three populations** of unlicensed OneDrive accounts, calculates exactly where each account is in the Day 60 / Day 93 timeline, and exports a prioritized CSV report.

---

## Features

- ✅ **No SPO PowerShell module required** — pure Microsoft Graph API
- ✅ **Multi-geo aware** — a single Graph token covers NAM, APC, CAN, DEU, GBR, IND, JPN automatically
- ✅ **Certificate or Client Secret authentication** — supports both auth flows
- ✅ **Detects all three unlicensed populations** — active users with no license + soft-deleted users + already-archived OneDrive sites
- ✅ **Finds sites Microsoft has already archived** — enumerates `GET /sites/getAllSites` and detects archived personal OneDrives via HTTP 423 response (optional, requires `Sites.Read.All`)
- ✅ **Audit log enrichment** — queries `directoryAudits` to find the exact date each license was removed (optional, requires `AuditLog.Read.All`)
- ✅ **Throttle handling** — exponential backoff with Retry-After support (429, 502, 503, 504)
- ✅ **Traffic-light urgency labels** — `CRITICAL`, `WARNING`, `MONITOR`, `OK`, `ARCHIVED`
- ✅ **UTF-8 BOM CSV output** — Excel-safe encoding
- ✅ **Email alert notifications** — sends HTML alert emails via Microsoft Graph API (`Mail.Send`) to a configurable list of admins/groups before sites go read-only or are archived. No SMTP relay required.

---

## Prerequisites

### PowerShell
- PowerShell 5.1 or later (Windows)

### Azure App Registration
A single app registration in the **home tenant** is required. The app must be granted the following **Application** permissions (not Delegated) with admin consent:

| Permission | Required | Purpose |
|---|---|---|
| `User.Read.All` | ✅ Required | Enumerate all users and read `assignedPlans` to detect unlicensed accounts |
| `Directory.Read.All` | ✅ Required | Read soft-deleted users from the Entra ID 30-day recycle bin |
| `Files.Read.All` | ✅ Required | Read OneDrive drive metadata for any user across all geo locations |
| `AuditLog.Read.All` | ⚠️ Optional | Query `directoryAudits` to find the exact license removal date. Set `$includeLicenseRemovalDates = $false` to skip. |
| `Sites.Read.All` | ⚠️ Optional | Enumerate all SharePoint sites via `GET /sites/getAllSites` to find personal OneDrive sites already archived by Microsoft. Set `$GetCurrentlyArchived = $false` to skip. |
| `Mail.Send` | ⚠️ Optional | Send HTML alert emails via Graph API (`POST /users/{sender}/sendMail`). Set `$SendEmailNotifications = $false` to skip. The `$EmailFrom` address must be a licensed Exchange Online mailbox. |

### Authentication
The script supports two authentication methods. Configure one in the `#region Configuration` block:

**Option A — Certificate (recommended)**
1. Generate a self-signed certificate or use an existing one
2. Upload the certificate public key to the app registration
3. Install the certificate (with private key) on the machine running the script
4. Set `$AuthType = 'Certificate'`, `$Thumbprint`, and `$CertStore`

**Option B — Client Secret**
1. Create a client secret in the app registration
2. Set `$AuthType = 'ClientSecret'` and `$clientSecret`

> ⚠️ Never commit a client secret to source control. Use environment variables or a secrets manager in production.

---

## Configuration

All configuration is in the `#region Configuration` block at the top of the script. Edit these values before running:

```powershell
# Tenant and App Registration
$tenantId  = 'your-tenant-id'
$clientId  = 'your-app-client-id'

# Auth: 'Certificate' or 'ClientSecret'
$AuthType  = 'Certificate'
$Thumbprint = 'YOUR_CERT_THUMBPRINT'
$CertStore  = 'LocalMachine'   # or 'CurrentUser'
$clientSecret = ''             # only used when $AuthType = 'ClientSecret'

# Output folder for the CSV report
$OutputFolder = $env:TEMP

# Microsoft's documented thresholds (do not change unless Microsoft updates them)
$ReadOnlyThresholdDays = 60
$ArchiveThresholdDays  = 93

# Audit log settings
$includeLicenseRemovalDates = $true   # Set $false to skip Phase 3 (faster)
$AuditLogLookbackDays       = 180     # Max supported lookback

# Archived OneDrive sites (Sites API)
$GetCurrentlyArchived = $true         # Set $false to skip Phase 2b (no Sites.Read.All needed)

# Throttle settings
$MaxRetries          = 15
$InitialBackoffSec   = 3
$RequestTimeoutSec   = 300

# Optional delay between OneDrive drive queries (seconds). 0 = no delay.
$delayBetweenRequests = 0

# Email notifications (requires Mail.Send Application permission)
$SendEmailNotifications    = $false          # Set $true to enable
$EmailTo                   = @(
    'admin@contoso.com'                       # Individual address or mail-enabled group
    'it-admins@contoso.com'
)
$EmailFrom                 = 'onedrive-alerts@contoso.com'   # Licensed Exchange Online mailbox
$DaysToNotifyBeforeReadOnly = 14             # Alert X days before read-only
$DaysToNotifyBeforeArchive  = 14             # Alert X days before archive
```

---

## Script Flow

The script executes in **5 phases** (plus Phase 2b) plus initialization and output steps:

```
┌─────────────────────────────────────────────────────────────────┐
│  AUTHENTICATION                                                  │
│  AcquireToken — Certificate JWT or Client Secret flow           │
│  Single Graph token covers all geo datacenters                  │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 1 — Active Unlicensed Users                              │
│  GET /users?$select=id,userPrincipalName,...,assignedPlans      │
│  Pages through ALL active Entra ID users ($top=999 per page)    │
│  Checks each user's assignedPlans for an enabled OneDrive or    │
│  SharePoint service plan ID                                     │
│  → Collects users with NO enabled plan into $activeUnlicensed   │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 2 — Soft-Deleted Users                                   │
│  GET /directory/deletedItems/microsoft.graph.user               │
│  Returns users in the Entra ID 30-day recycle bin               │
│  Uses deletedDateTime as the UnlicensedDate                     │
│  → Collects into $softDeletedUsers                              │
│  NOTE: Users deleted >30 days ago are purged — not discoverable │
│        by Phases 1 or 2; use Phase 2b instead                   │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  MERGE — De-duplicate                                           │
│  Combines Phase 1 + Phase 2 populations                         │
│  Removes any user appearing in both sets (dedup on UserId)      │
│  → $allCandidates                                               │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 2b — Already-Archived OneDrive Sites (optional)          │
│  Requires: $GetCurrentlyArchived = $true + Sites.Read.All       │
│                                                                 │
│  Step A: GET /sites/getAllSites — enumerate all tenant sites     │
│          Filter: isPersonalSite = true                          │
│          → $personalSites list                                  │
│                                                                 │
│  Step B: For each personal site:                                │
│          GET /beta/sites/{id}?$select=id,siteCollection         │
│          HTTP 423 Locked = site is archived by Microsoft        │
│          (Graph refuses metadata requests for archived sites)   │
│          200 OK with null archivalDetails = active site, skip  │
│                                                                 │
│  Deduplicates against Phase 1/2 results by UPN                 │
│  → $archivedSites (UserSource = 'Archived')                     │
│  Skipped if $GetCurrentlyArchived = $false                      │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 3 — Audit Log Enrichment (optional)                      │
│  GET /auditLogs/directoryAudits                                 │
│  Queries two activity types in sequence:                        │
│    1. "Change user license"                                     │
│    2. "Remove user from licensed group"                         │
│  Filtered by $AuditLogLookbackDays in PowerShell after fetch    │
│  (API-side activityDateTime filter causes 400 in some tenants)  │
│  Builds userId → most-recent-event-date lookup table            │
│  → Populates UnlicensedDate on $activeUnlicensed entries        │
│  Skipped if $includeLicenseRemovalDates = $false                │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 4 — OneDrive Drive Query                                 │
│  GET /users/{id}/drive  (one call per Phase 1/2 candidate)      │
│  Graph routes each call to the correct geo automatically        │
│  404 = no OneDrive exists → skipped from report                │
│  403 / timeout = access error → included in report for review  │
│  → $confirmedUnlicensed                                         │
│  Phase 2b sites are merged in AFTER Phase 4 (drive already     │
│  queried inside Get-ArchivedOneDriveSites — 423 = no data)      │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 5 — Milestone Calculations                               │
│  For Phase 1/2 accounts with a known UnlicensedDate:            │
│    ReadOnlyDate        = UnlicensedDate + 60 days               │
│    ArchiveDate         = UnlicensedDate + 93 days               │
│    DaysSinceUnlicensed = Today - UnlicensedDate                 │
│    DaysUntilReadOnly   = ReadOnlyDate - Today                   │
│    DaysUntilArchive    = ArchiveDate - Today                    │
│  For Phase 2b accounts (no UnlicensedDate available):           │
│    UrgencyStatus = ARCHIVED - Currently Archived                │
│  Assigns UrgencyStatus traffic-light label (see below)          │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  OUTPUT                                                         │
│  Sorts report: CRITICAL → ARCHIVED → WARNING → MONITOR → OK    │
│  Exports UTF-8 BOM CSV to $OutputFolder                         │
│  Prints summary to console (by source and by urgency)           │
│  File: UnlicensedOneDrive_<yyyyMMddHHmmss>.csv                 │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 10 — Email Notifications (optional)                       │
│  Skipped if $SendEmailNotifications = $false                    │
│                                                                 │
│  Two alerts fired independently:                                │
│    (1) Read-Only alert — accounts where DaysUntilReadOnly       │
│        is >= 0 and <= $DaysToNotifyBeforeReadOnly               │
│    (2) Archive alert  — accounts where DaysUntilArchive         │
│        is >= 0 and <= $DaysToNotifyBeforeArchive                │
│                                                                 │
│  Each alert sends one HTML email via Graph API:                 │
│    POST /users/{EmailFrom}/sendMail                             │
│  Uses the existing bearer token — no SMTP relay required        │
│  Supports individual mailboxes and mail-enabled groups in       │
│  $EmailTo                                                       │
│  Requires Mail.Send (Application) on the app registration       │
└─────────────────────────────────────────────────────────────────┘
```

### Urgency Status Labels

| Label | Source | Condition |
|---|---|---|
| `CRITICAL - Archives TODAY` | Phase 1/2 | ArchiveDate = today |
| `CRITICAL - Archives within 7 days` | Phase 1/2 | DaysUntilArchive ≤ 7 |
| `ARCHIVED - Past Day 93` | Phase 1/2 | DaysUntilArchive < 0 (date known) |
| `ARCHIVED - Currently Archived` | Phase 2b | Site returned HTTP 423 Locked |
| `WARNING - Goes Read-Only TODAY` | Phase 1/2 | DaysUntilReadOnly = today |
| `WARNING - Read-Only within 7 days` | Phase 1/2 | DaysUntilReadOnly ≤ 7 |
| `WARNING - Read-Only, Archive pending` | Phase 1/2 | Past Day 60, archive > 7 days away |
| `MONITOR - Read-Only within 30 days` | Phase 1/2 | DaysUntilReadOnly ≤ 30 |
| `OK - More than 30 days remaining` | Phase 1/2 | DaysUntilReadOnly > 30 |
| `Unknown - No Unlicensed Date` | Phase 1 | No audit event found within lookback window |

---

## Usage

```powershell
# Edit the configuration section, then run:
.\Get-UnlicensedOneDriveReport.ps1
```

The script is entirely self-contained — no parameters, no module imports. All settings are in the configuration block at the top.

---

## Output — CSV Columns

| Column | Description |
|---|---|
| `UserSource` | `Active` (license removed), `SoftDeleted` (deleted from Entra, within 30-day recycle bin), or `Archived` (OneDrive already archived by Microsoft — Entra user purged >30 days ago) |
| `DisplayName` | User's display name |
| `UserPrincipalName` | UPN |
| `AccountEnabled` | Whether the Entra account is still enabled |
| `UnlicensedDueTo` | `License removed by admin`, `Owner deleted from Entra ID`, or `OneDrive archived by Microsoft` |
| `UnlicensedDate` | Date/time the license was removed (from audit log or deletedDateTime) |
| `DaysSinceUnlicensed` | Days elapsed since UnlicensedDate |
| `ReadOnlyDate` | Date OneDrive goes read-only (Day 60) |
| `ArchiveDate` | Date OneDrive is archived/deleted (Day 93) |
| `DaysUntilReadOnly` | Days remaining until read-only (negative = already read-only) |
| `DaysUntilArchive` | Days remaining until archived (negative = already archived) |
| `UrgencyStatus` | Traffic-light label (see table above) |
| `StorageUsedGB` | OneDrive storage used (GB) |
| `StorageTotalGB` | OneDrive storage quota (GB) |
| `DriveUrl` | Direct URL to the OneDrive |
| `DriveLastModified` | Last modification timestamp of the drive |
| `DriveType` | Drive type (e.g., `business`) |
| `Notes` | For Phase 2b archived sites: `Site is archived (HTTP 423 Locked) — storage details unavailable while archived`. For Phase 1/2 errors: HTTP status and message. |

---

## Detected Service Plans

The script detects a user as unlicensed when **none** of the following service plan IDs appear in their `assignedPlans` with `capabilityStatus = 'Enabled'`:

### OneDrive Plans
| Plan Name | ID |
|---|---|
| ONEDRIVELITE_IW (Office for the web with OneDrive, Basic Collaboration) | `b4ac11a0` |
| ONEDRIVECLIPCHAMP (OneDrive with Clipchamp Premium) | `f7e5b77d` |
| ONEDRIVESTANDARD (OneDrive Plan 1 — M365 Apps, E1) | `13696edf` |
| ONEDRIVE_BASIC variant | `4495894f` |
| ONEDRIVEENTERPRISE (OneDrive Plan 2 — standalone) | `afcafa6a` |
| ONEDRIVE_BASIC (Visio plans) | `da792a53` |
| ONEDRIVE_BASIC_GOV (Government) | `98709c2e` |

### SharePoint Plans
| Plan Name | ID |
|---|---|
| SHAREPOINTWAC (Office for the web — E1/E3/E5) | `e95bec33` |
| SHAREPOINTENTERPRISE (SharePoint Plan 2 — E3/E5) | `5dbe027f` |
| SHAREPOINTDESKLESS (F-tier/Teams Free) | `902b47e5` |
| SHAREPOINTENTERPRISE_EDU (Education) | `63038b2c` |
| SHAREPOINTENTERPRISE_MIDMARKET | `6b5b6a67` |
| SHAREPOINTSTANDARD (SharePoint Plan 1 — standalone) | `c7699d2e` |
| SHAREPOINTSTANDARD_EDU (Education) | `0a4983bb` |

This covers **O365/M365 E1, E3, E5**, Business Basic/Standard/Premium, F1/F3, and Education SKUs.

---

## Limitations

| Limitation | Detail |
|---|---|
| Users deleted > 30 days ago | Permanently purged from the Entra ID recycle bin — not visible in Phases 1 or 2. However, **Phase 2b** (`$GetCurrentlyArchived = $true`) can still find these accounts as long as Microsoft has archived (not yet deleted) their OneDrive. Graph returns HTTP 423 Locked for archived personal sites, which the script detects. Once Microsoft purges the OneDrive entirely, the site disappears from `getAllSites` and cannot be found by any Graph call. |
| Audit log lookback | `directoryAudits` retains data for a maximum of **180 days**. License changes older than `$AuditLogLookbackDays` will show `Unknown - No Unlicensed Date`. |
| Audit log API date filter | The `activityDateTime ge` OData filter causes HTTP 400 in some tenants. The script works around this by fetching all events for each activity type and filtering by date in PowerShell. |
| Guest / external users | Guest accounts without a OneDrive license are included in the scan but typically return 404 on the drive query and are filtered out. |
| Phase 2b storage data | Archived sites return HTTP 423 Locked on both the site metadata and drive queries, so `StorageUsedGB` and `StorageTotalGB` are blank for `Archived` source entries. Use the SharePoint admin center for exact storage figures. |
| Phase 2b UPN reconstruction | UPNs are reconstructed from the personal site URL (e.g. `john_doe_contoso_com` → `john.doe@contoso.com`). Usernames containing `.` or `_` are ambiguous after SharePoint's encoding and the reconstructed UPN may differ from the original. `DisplayName` is always accurate. |

---

## Throttling & Performance

The script includes a full throttle-handling wrapper (`Invoke-GraphRequestWithThrottleHandling`) that:

- Respects `Retry-After` headers on 429 responses
- Applies exponential backoff for 502/503/504 and network timeouts
- Retries up to `$MaxRetries` (default: 15) times per request
- Caps backoff at 300 seconds

For large tenants (10,000+ users), Phase 4 (drive queries) is the most time-consuming step as it makes one Graph call per candidate. Set `$delayBetweenRequests = 0` for maximum speed, or increase it if you need to reduce API load.

Phase 2b makes **two Graph calls per personal OneDrive site** (one individual site detail GET to detect 423). In a tenant with hundreds of personal sites this adds a modest number of extra calls but is bounded by the number of personal OneDrives, not total users.

---

## Email Notifications

When `$SendEmailNotifications = $true`, the script sends up to two HTML alert emails after the report is generated — one for accounts approaching read-only and one for accounts approaching archive. Emails are delivered via the **Microsoft Graph API** (`POST /users/{sender}/sendMail`) using the same bearer token already acquired for data collection. No SMTP relay, no credentials, and no deprecated `Send-MailMessage` cmdlet are used.

### Configuration

| Variable | Default | Description |
|---|---|---|
| `$SendEmailNotifications` | `$false` | Master on/off toggle. Set to `$true` to enable. |
| `$EmailTo` | *(array)* | One or more recipient addresses. Accepts individual mailboxes and mail-enabled groups/distribution lists. |
| `$EmailFrom` | *(string)* | Sender address. Must be a **licensed Exchange Online mailbox** in the tenant. |
| `$DaysToNotifyBeforeReadOnly` | `14` | Send the read-only alert when `DaysUntilReadOnly` is at or below this value. Set to `0` to alert only on the day of the event. |
| `$DaysToNotifyBeforeArchive` | `14` | Send the archive alert when `DaysUntilArchive` is at or below this value. |

### Email content

Each alert email contains:
- A colour-coded HTML table of affected accounts (red ≤ 3 days remaining, amber ≤ 7 days, white otherwise)
- Columns: Display Name, UPN, Source, Storage Used, target date, days remaining, urgency status
- Tenant ID, report run date, and the relevant Day threshold in the header
- Path to the exported CSV report

### Required permission

Add `Mail.Send` (**Application** type, not Delegated) to the app registration and grant admin consent:

1. **Entra admin center** → App registrations → your app → **API permissions**
2. **Add a permission** → Microsoft Graph → Application permissions → `Mail.Send`
3. **Grant admin consent** for the tenant

> ⚠️ `$EmailFrom` must be a user with an active Exchange Online mailbox (E1/E3/E5, Exchange Online Plan 1/2, or equivalent). A cloud-only account without an Exchange license will return HTTP 403.

---

## Security Notes

- **Certificate authentication is recommended** over client secrets for production use
- The script uses `GetRSAPrivateKey()` (supports both CAPI and CNG/KSP certificate providers) rather than the legacy `.PrivateKey` property
- Do not store client secrets in the script file — use environment variables or a secrets manager
- The core app registration requires only **read** permissions — no write access to any resource
- `Mail.Send` (Application) is the only write-capable permission and is entirely optional — set `$SendEmailNotifications = $false` to run without it
- Email is sent via Graph API using the existing access token — no SMTP credentials are stored or transmitted

---

## References

- [Unlicensed OneDrive accounts — Microsoft docs](https://learn.microsoft.com/en-us/sharepoint/unlicensed-onedrive-accounts)
- [Graph API: List users](https://learn.microsoft.com/en-us/graph/api/user-list)
- [Graph API: Get drive](https://learn.microsoft.com/en-us/graph/api/drive-get)
- [Graph API: List directoryAudits](https://learn.microsoft.com/en-us/graph/api/directoryaudit-list)
- [Graph API: List all sites (getAllSites)](https://learn.microsoft.com/en-us/graph/api/site-getallsites)
- [Graph API: siteArchivalDetails resource](https://learn.microsoft.com/en-us/graph/api/resources/sitearchivaldetails)
- [Graph API: Send mail](https://learn.microsoft.com/en-us/graph/api/user-sendmail)
- [Service plan identifiers for licensing](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference)

---

## Author

**Mike Lee**  
Created: April 28, 2026
