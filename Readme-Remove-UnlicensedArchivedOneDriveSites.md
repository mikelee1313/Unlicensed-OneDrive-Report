# Remove Unlicensed Archived OneDrive Sites

A PowerShell administration script for reviewing and deleting selected OneDrive sites in SharePoint Online.

[Remove-UnlicensedArchivedOneDriveSites.ps1](./Remove-UnlicensedArchivedOneDriveSites.ps1) supports two input modes:

- **Report mode:** process candidates from a SharePoint Admin Center unlicensed OneDrive CSV export, applying explicit eligibility filters.
- **Explicit-list mode:** process administrator-selected URLs from `$OneDriveSiteList`, without report-based eligibility filtering.

Site lookup, unlocking, and deletion use **PnP.PowerShell**. The existing report download helper retains its REST-based implementation.

> **Destructive operation:** setting `$PerformDeletes = $true` enables deletion without another switch or confirmation prompt. Start with a dry run, review the audit, and obtain the appropriate authorization before enabling deletion.
>
> **Retention is not bypassed:** an unlocked site can still be blocked by retention or holds. Explicit-list mode does not override those service protections.

## Contents

- [Features and scope](#features-and-scope)
- [Prerequisites](#prerequisites)
- [Authentication setup](#authentication-setup)
- [Quick start](#quick-start)
- [Configuration reference](#configuration-reference)
- [Input modes and precedence](#input-modes-and-precedence)
- [CSV format and eligibility](#csv-format-and-eligibility)
- [Processing workflow](#processing-workflow)
- [Multi-geo tenants](#multi-geo-tenants)
- [Vanity domains](#vanity-domains)
- [Audit journal](#audit-journal)
- [Errors and automation](#errors-and-automation)
- [Report download helper](#report-download-helper)
- [Troubleshooting](#troubleshooting)
- [Limitations and operational guidance](#limitations-and-operational-guidance)
- [Publishing checklist](#publishing-checklist)
- [References](#references)

## Features and scope

- Dry-run mode performs live site lock-state checks without unlocking or deleting sites.
- CSV mode requires recognized archive, unlicensed, and deletion-block fields.
- Duplicate CSV URLs are normalized and processed once.
- Conflicting eligibility values for the same CSV URL fail closed.
- Explicit site lists override CSV input and automatic discovery.
- Personal-site URLs must match exactly one configured SharePoint admin host.
- Locked sites can be unlocked and rechecked before deletion.
- Deletion moves sites to **Deleted Sites**; the script does not permanently purge them.
- Every run uses a unique, append-only CSV audit journal.
- Mutation-intent checkpoints are saved before unlock and delete calls.
- Audit-write failures stop processing before further site operations.
- Handled site failures are audited, and later sites can continue.
- Named PowerShell regions organize configuration, authentication, operations, and execution.

The script is not a general SharePoint site cleanup tool. URL validation restricts targets to OneDrive personal-site roots.

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | PowerShell 7.2 or later, subject to the selected PnP module's own runtime requirements. Use `pwsh`, not Windows PowerShell 5.1. |
| PnP.PowerShell | The script's documented dependency is PnP.PowerShell 2.x. Version 2.12.0 was used for the read-only integration check. Newer major versions are not certified by these checks. |
| Host | The certificate implementation uses the Windows `Cert:` provider. Windows is the documented host for this configuration. |
| Microsoft Entra application | An app registration in the tenant, with the required SharePoint application permission and tenant admin consent. |
| Permission | **SharePoint > Application permissions > `Sites.FullControl.All`**. This is a broad permission; protect the application and its credentials accordingly. |
| Certificate | An RSA certificate registered with the application, with its private key accessible to the Windows account running the script. |
| Connectivity | Access to Microsoft Entra authentication and the configured SharePoint Online admin endpoints. |
| Input | A current, compatible report CSV or an explicitly approved site list. |
| Audit destination | A writable output folder with adequate disk space. |

Example installation for the tested module version:

```powershell
Install-Module PnP.PowerShell -RequiredVersion 2.12.0 -Scope CurrentUser
Import-Module PnP.PowerShell -RequiredVersion 2.12.0
```

The script imports PnP.PowerShell on demand but does not pin a version internally. Explicitly import your validated version if multiple versions are installed.

## Authentication setup

### Certificate authentication: recommended

1. Create or use an approved Microsoft Entra app registration.
2. Add the **SharePoint application permission** `Sites.FullControl.All`.
3. Obtain tenant-wide admin consent.
4. Upload the certificate's public portion to the app registration.
5. Install the certificate with its private key in the executing machine's `CurrentUser\My` or `LocalMachine\My` certificate store.
6. Ensure the executing account can access the private key. For scheduled jobs, check the job account rather than only your interactive account.
7. Edit the script's Configuration section:

```powershell
$tenantId = '<tenant-id>'
$clientId = '<application-client-id>'
$AuthType = 'Certificate'
$Thumbprint = '<certificate-thumbprint>'
$CertStore = 'CurrentUser' # Or LocalMachine
$clientSecret = ''
```

The script constructs a certificate-signed client assertion, requests a SharePoint-scoped client-credentials token, and supplies that token to `Connect-PnPOnline -AccessToken`.

Tokens and PnP connections are cached separately by admin URL. Expiring tokens are refreshed, and connections are rebuilt when their token changes.

### Client-secret configuration

The script also retains a `$AuthType = 'ClientSecret'` token-acquisition branch. Its presence does **not** establish that secret-based app-only authentication is supported for every SharePoint/PnP operation. That path has not been validated end-to-end here.

Use certificate authentication for the documented workflow. Do not commit client secrets, access tokens, certificate private keys, or PFX passwords to GitHub.

## Quick start

### 1. Configure the tenant and output folder

Edit the Configuration region in [the script](./Remove-UnlicensedArchivedOneDriveSites.ps1):

```powershell
$SPOAdminUrls = @(
    'https://contoso-admin.sharepoint.com'
)

$OutputFolder = 'C:\OneDriveCleanup'
$OneDriveSiteList = @()
$PerformDeletes = $false
$UnlockLockedSitesBeforeDelete = $true
```

Create the output folder before using automatic discovery, and place a current report there:

```powershell
New-Item -ItemType Directory -Path 'C:\OneDriveCleanup' -Force
```

### 2. Run a report-based dry run

From the folder containing the script:

```powershell
.\Remove-UnlicensedArchivedOneDriveSites.ps1 `
    -CsvPath 'C:\OneDriveCleanup\UnlicensedOneDrive_report.csv'
```

Using an explicit CSV path is recommended for predictable targeting. Review the printed summary and the audit journal.

### 3. Enable deletion only after review

Change this value **inside the script**:

```powershell
$PerformDeletes = $true
```

Run the same command again. There is no additional `-DeleteSites` switch.

When finished, set `$PerformDeletes` back to `$false`.

> The Configuration section assigns these variables on every invocation. Setting `$PerformDeletes` in the caller's shell does not override the value assigned inside the script.

### Alternative: administrator-provided site list

Set:

```powershell
$OneDriveSiteList = @(
    'https://contoso-my.sharepoint.com/personal/alex_contoso_com'
    'https://contoso-my.sharepoint.com/personal/jamie_contoso_com'
)

$PerformDeletes = $false
```

Then run:

```powershell
.\Remove-UnlicensedArchivedOneDriveSites.ps1
```

Review the dry run before changing `$PerformDeletes` to `$true`.

**Explicit-list mode bypasses CSV archive, licensing, and deletion-block checks.** The administrator must independently establish that each site is appropriate for deletion. SharePoint still enforces retention and holds.

## Configuration reference

Except for `-CsvPath`, settings are edited inline in the script.

| Setting | Default or expected value | Purpose |
|---|---|---|
| `$tenantId` | Your tenant GUID | Tenant used for token acquisition. |
| `$clientId` | Your application GUID | App registration client ID. |
| `$AuthType` | `'Certificate'` | Authentication branch; certificate recommended. |
| `$Thumbprint` | Your certificate thumbprint | Certificate lookup in the configured store. |
| `$CertStore` | `'LocalMachine'` | Use `LocalMachine` or `CurrentUser`. |
| `$clientSecret` | `''` | Retained secret-authentication setting; never publish a real secret. |
| `$SPOAdminUrls` | Tenant-specific array | One SharePoint admin URL per geo, without a trailing slash. |
| `$SPOAdminUrlMappings` | `@{}` | Explicit admin-to-OneDrive-host mapping for vanity domains; optional `TokenResourceUrl` overrides the OAuth resource origin. |
| `$OutputFolder` | `$env:TEMP` | Automatic report discovery and audit output location. A dedicated folder is recommended. |
| `$OneDriveSiteList` | `@()` | Nonempty array overrides both `-CsvPath` and automatic discovery. |
| `$PerformDeletes` | `$false` | `$false`: dry run. `$true`: permit unlock/delete calls without another switch. |
| `$UnlockLockedSitesBeforeDelete` | `$true` | Permit unlocking eligible locked sites in delete mode. |
| `$debug` | `$false` | Additional request logging in the report REST wrapper. |
| `$MaxRetries` | `15` | Retry limit for the report REST wrapper, not a custom PnP deletion retry policy. |
| `$InitialBackoffSec` | `3` | Initial report REST retry delay. |
| `$RequestTimeoutSec` | `300` | Timeout for individual report REST requests. |
| `$SPOExportPollIntervalSec` | `5` | Report helper polling interval. |
| `$SPOExportMaxWaitSec` | `120` | Report helper polling wait budget; request/retry time can extend total elapsed time. |
| `-CsvPath` | Not supplied | Select one report explicitly, unless the site list is populated. |

The defaults describe the intended safe template. Check your local configuration before each run.

## Input modes and precedence

The script selects input in this order:

1. **Nonempty `$OneDriveSiteList`:** validate, normalize, and deduplicate the supplied URLs. Ignore `-CsvPath` and report discovery.
2. **`-CsvPath`:** import the specified report.
3. **Automatic discovery:** inspect `$OutputFolder` for `UnlicensedOneDrive*.csv`.

Automatic discovery:

- Excludes files named `UnlicensedOneDrive_DeleteAudit_*`.
- Sorts by `LastWriteTime`, newest first.
- Checks candidates until it finds a nonempty report with the required columns.
- Warns and continues when a candidate is unreadable or incompatible.
- Selects **one** report, not every report in the folder.

Discovery uses file modification time, not the report's business-data timestamp, and does not impose a report-age limit. Supply `-CsvPath` when the exact input matters.

## CSV format and eligibility

### Required columns

At least one recognized header is required for each category:

| Category | Accepted header names |
|---|---|
| Site URL | `URL`, `SiteUrl` |
| Archive status | `ARCHIVE_STATUS`, `ArchiveStatus`, `Archive status` |
| Unlicensed reason | `UNLICENSED_REASON`, `Unlicensed due to`, `UnlicensedDueTo` |
| Deletion blocker | `DELETION_BLOCK_REASON`, `Deletion blocked by`, `DeletionBlockedBy` |

An empty report or missing required column category causes a run-level failure. If several aliases for a category are present, the first matching alias in the table's order is used.

### Eligible values

CSV rows must satisfy **all** of these conditions:

| Field | Allowed values |
|---|---|
| Archive status | `Archived`, `RecentlyArchived`, `FullyArchived` |
| Unlicensed reason | `Owner deleted from Entra ID`, `License removed by admin` |
| Deletion blocker | Blank, `0`, `None`, `Not blocked`, `No blockers` |

Values are trimmed, and comparisons use PowerShell's default case-insensitive behavior. Unrecognized status/reason values and other blocker values are skipped with an explanatory audit reason.

Minimal example:

```csv
URL,Archive status,Unlicensed due to,Deletion blocked by
https://contoso-my.sharepoint.com/personal/alex_contoso_com,Archived,License removed by admin,
https://contoso-my.sharepoint.com/personal/jamie_contoso_com,Archived,Owner deleted from Entra ID,Retention policy
```

The first row is a candidate for a live lock-state check. The second is skipped because the report identifies a retention blocker.

### Duplicate handling

CSV mode trims URLs and normalizes parseable personal-site URLs, including removal of a trailing slash. Duplicate URL groups are evaluated before site operations.

- Matching eligibility values: process the site once.
- Different archive, unlicensed, or deletion-block values: record `Failed` and do not query, unlock, or delete that site.

Conflicting rows are not resolved by taking the newest row or preferring an unblocked row. Even different allowed spellings such as a blank blocker and `None` are treated as conflicting values. Resolve the source discrepancy or export a fresh report before retrying.

### URL validation

Targets must:

- Be absolute HTTPS URLs on the default HTTPS port.
- Have a path of `/personal/<site>` with an optional trailing slash.
- Have no credentials, query string, or fragment.
- Match exactly one configured admin host.

Use the personal-site root, not a document-library, file, sharing, or subfolder URL.

Invalid explicit-list URLs fail input selection before the audit is initialized. Invalid CSV URLs are recorded as site failures during processing.

## Processing workflow

1. Select and validate input.
2. Create a uniquely named audit file with its CSV header, without overwriting an existing file.
3. In CSV mode, normalize/deduplicate rows and evaluate report eligibility.
4. Route each target to the matching configured geo.
5. Query the site using `Get-PnPTenantSite -Detailed`.
6. Require a recognized lock state: `Unlock`, `ReadOnly`, `NoAccess`, or `NoAdditions`.
7. In dry-run mode, record the intended action without changing the site.
8. In delete mode, optionally write an unlock checkpoint and call `Set-PnPTenantSite -LockState Unlock -Wait`.
9. Wait five seconds after an unlock attempt and query the lock state again.
10. Require `Unlock` before deletion.
11. Save a deletion-intent checkpoint and call `Remove-PnPTenantSite -Force`.
12. Append the final outcome to the audit and display the summary.

If unlocking is disabled, locked sites are skipped. Unknown lock states or a failed unlock verification produce a site failure.

The script does not restore the original lock state if deletion fails after a successful unlock. Review such sites after the run.

## Multi-geo tenants

Configure the actual admin URLs for every geo containing targets:

```powershell
$SPOAdminUrls = @(
    'https://contoso-admin.sharepoint.com'
    'https://contosoEUR-admin.sharepoint.com'
)
```

These are illustrative URLs; use your tenant's real geo hostnames.

For standard SharePoint admin domains, routing derives the personal-site host by replacing `-admin.` with `-my.`. An explicit mapping overrides that derived host list. Each target must have exactly one matching admin entry. The same app registration is used, with tokens and connections cached per admin URL.

A selected CSV can include targets across configured geos. Automatic discovery still selects only one report; it does not merge separate geo exports. Run separate explicitly selected reports when necessary.

## Vanity domains

An administrator-facing vanity hostname need not follow the standard SharePoint naming pattern. Do not infer the OneDrive hostname from it. Configure the actual admin endpoint in `$SPOAdminUrls`, then map its exact OneDrive hostnames in `$SPOAdminUrlMappings`.

Generic example:

```powershell
$SPOAdminUrls = @(
    'https://portal-admin.example.com'
)

$SPOAdminUrlMappings = @{
    'https://portal-admin.example.com' = @{
        OneDriveHosts = @(
            'personal.example.com'
            'contoso-my.sharepoint.com'
        )
        TokenResourceUrl = 'https://contoso-admin.sharepoint.com'
    }
}

$OneDriveSiteList = @(
    'https://personal.example.com/personal/alex_contoso_com'
)
$PerformDeletes = $false
```

Replace **all** example values with verified tenant values:

- Mapping keys must correspond to `$SPOAdminUrls` entries. HTTPS origins are normalized for case and trailing slash.
- `OneDriveHosts` contains exact DNS hostnames, **not URLs**, paths, or wildcard patterns. Configure only hostnames belonging to the intended tenant/geo.
- A mapping replaces the automatically derived host list for that admin URL. Include the standard OneDrive hostname too if targets can use both names.
- A target must match exactly one configured admin URL. Overlapping host mappings fail closed when a target is routed.
- Nonstandard admin domains require a mapping. Standard `-admin.sharepoint.com`, `.us`, `.de`, and `.cn` domains retain automatic routing when no mapping is supplied.
- Both CSV mode and explicit-list mode use the same mapping and retain the HTTPS personal-site-root validation.

### OAuth resource versus admin endpoint

`TokenResourceUrl` is optional. When omitted, the script requests `<admin-origin>/.default`, preserving its previous authentication behavior.

When the vanity endpoint is not the SharePoint resource registered with Microsoft Entra, set `TokenResourceUrl` to the **verified SharePoint OAuth resource origin** for that tenant/geo. Do not assume a DNS alias or browser redirect is a registered OAuth resource. Supply an HTTPS origin without a path, query string, custom port, or `/.default` suffix.

The override changes token acquisition only:

- `Connect-PnPOnline` and report requests still use the configured admin endpoint.
- The same tenant ID, app registration, and certificate are used.
- Tokens/connections remain cached by admin endpoint.
- The report download function's body is unchanged; its existing shared token helper uses this resource setting.

Routing support does not provision a custom domain, configure DNS/proxies, rewrite site identities, or prove that a vanity endpoint accepts PnP/CSOM requests. A browser-only vanity redirect may not work for those operations. Verify the endpoint and token audience with the tenant administrator and perform a read-only dry run before enabling deletion. If the vanity endpoint is not API-capable, use the tenant's actual API-capable SharePoint admin URL and map the accepted site hostnames to it.

These configuration paths have offline mocked coverage; live vanity-domain authentication and operations have not been validated against a customer tenant.

## Audit journal

### Filename and persistence

Audit files are saved under `$OutputFolder`:

```text
UnlicensedOneDrive_DeleteAudit_<yyyyMMddHHmmss>_<run-guid>.csv
```

The header is written before site processing. Each checkpoint and final outcome is appended as processing proceeds. Audit initialization failure prevents site mutations; a later write failure immediately stops further operations.

This is an operational journal, not a transactional rollback mechanism or a tamper-proof compliance log.

### Columns

| Column | Meaning |
|---|---|
| `SiteUrl` | Target personal-site URL. |
| `AdminUrl` | Selected geo admin URL; may be blank if validation failed earlier. |
| `LockStateBefore` | Lock state observed before processing. |
| `LockStateAfter` | Last lock state recorded by this workflow; not a post-deletion query. |
| `UnlockAttempted` | Unlock intent/attempt flag; see checkpoint caveat below. |
| `DeleteAttempted` | Delete intent/attempt flag; see checkpoint caveat below. |
| `DeleteSucceeded` | The deletion command returned successfully. |
| `Outcome` | `InProgress`, `Deleted`, `DryRun`, `Skipped`, or `Failed`. |
| `Reason` | Action description, skip explanation, or error message. |
| `CSV` | Input CSV path or `Configuration: OneDriveSiteList`. |

### Outcomes

| Outcome | Interpretation |
|---|---|
| `InProgress` | A mutation checkpoint was persisted; completion has not yet been recorded. |
| `Deleted` | The remove command returned successfully; the script reports the site moved to Deleted Sites. |
| `DryRun` | No mutation occurred; `Reason` describes the intended action. |
| `Skipped` | Filtering or lock-state rules prevented an attempt. |
| `Failed` | A site operation or validation failed; inspect `Reason`. |

A site can have multiple journal rows: for example, an unlock checkpoint, a deletion checkpoint, and a final result. **Do not count every CSV row as a separate site or deletion.** The console summary uses final outcomes collected in memory.

Attempt flags are set before the mutation call so intent can be recorded first. An `InProgress` row with `DeleteAttempted = True` is **not proof** that SharePoint received the request.

If the last row for a site is still `InProgress`, inspect the current live site/deleted-site state before retrying. A process interruption or failed final audit write can leave an unconfirmed checkpoint even if the service operation completed.

### Review the latest row for each site

For one audit file:

```powershell
$auditPath = 'C:\OneDriveCleanup\UnlicensedOneDrive_DeleteAudit_<timestamp>_<run-guid>.csv'
$journal = @(Import-Csv -LiteralPath $auditPath)

$latest = @(
    $journal |
        Group-Object SiteUrl |
        ForEach-Object { $_.Group[-1] }
)

$latest | Select-Object SiteUrl, Outcome, Reason | Format-Table -Wrap
$latest | Group-Object Outcome | Select-Object Name, Count
```

Do not filter out `InProgress` before selecting the last row: an incomplete operation must remain visible.

CSV boolean values import as strings. Compare `DeleteSucceeded` with `'True'`, or explicitly parse it; casting the string `'False'` to `[bool]` does not produce the expected result.

## Errors and automation

### Handled site failures

Retention blocks, permission errors raised during site operations, failed unlock verification, and other per-site errors are caught and audited as `Failed`. Processing can continue to later sites.

If only handled site failures occur, the final summary is followed by a **warning**, not an intentional terminating exception.

Example:

```text
Complete. 1 row(s) reviewed: 0 deleted, 0 dry-run, 0 skipped, 1 failed; 0 run-level error(s).
WARNING: OneDrive processing completed with 1 failed site(s). These failures were handled and audited.
```

Retention remains enforced. The warning does not mean deletion succeeded.

### Terminating failures

The script still terminates for conditions such as:

- Invalid explicit-list input or no usable input.
- Report import/schema failures at run level.
- Audit initialization failure.
- Any audit checkpoint/final-row write failure.
- Unexpected run-level processing failures.

Audit-write failures stop immediately. Other captured run-level failures are reported in the final summary and then cause a terminating error.

### Automation implications

There is no explicit exit-code contract or structured top-level result object. **Do not use the absence of an exception as proof that all sites were deleted.** Automated callers must review the journal's latest outcome per site and treat `Failed` or unresolved `InProgress` entries according to their operational policy.

The script also has no native `-WhatIf`, `-Confirm`, `-DeleteSites`, or command-line `$PerformDeletes` parameter. Its dry-run/deletion control is the inline configuration value.

## Report download helper

`Get-UnlicensedOneDriveReport` remains in the script for the existing report export workflow. It:

1. Ensures a valid SharePoint token.
2. Obtains a form digest.
3. Requests an export through `/_api/SPO.Tenant/ExportToCSV`.
4. Polls for the generated file.
5. Downloads it to a tenant-labelled local CSV.

The report REST wrapper provides retry/backoff handling for selected transient responses. Site lookup, unlocking, and deletion do **not** use `/_api/SPO.Tenant` operational endpoints.

**Normal script execution does not call this helper automatically.** Export a current report before running CSV mode. The helper uses an existing admin-report REST implementation; do not interpret it as a guarantee of a stable public API contract.

The script is not a function-only module. Dot-sourcing it also executes its main workflow, including deletion if enabled. Do not dot-source it merely to load the download helper without accounting for that behavior.

## Troubleshooting

| Symptom | Explanation / recommended action |
|---|---|
| Site is blocked by retention | The service rejected deletion. Check applicable retention policies and holds with the responsible compliance administrator. Excluding a site from one policy does not prove all restrictions have cleared. Do not bypass compliance requirements. |
| Lock state is `Unlock`, but deletion fails | Lock state and retention eligibility are separate. Read the deletion error in the audit. |
| Dry run says deletion would be attempted, but real deletion fails | Dry run checks lock state and available report fields; it does not invoke deletion or fully discover live retention/hold restrictions. |
| Site was unlocked but not deleted | A later check or deletion failed. Inspect the site and decide whether its original lock state should be restored. The script does not roll it back. |
| PnP.PowerShell is missing | Install/import a validated module version in the same PowerShell environment used to run the script. |
| Certificate is not found | Verify thumbprint and `CertStore`, and check the executing account's certificate store. |
| RSA private key cannot be accessed | Ensure the certificate includes its RSA private key and the executing account has access. |
| Unauthorized / access denied | Check tenant/app IDs, certificate registration, SharePoint application permission, admin consent, and target geo. Review the exact error; do not assume authentication alone grants deletion rights. |
| No compatible report found | Check folder existence, filename pattern, required headers, and data rows. Use `-CsvPath` to select the intended report directly. |
| `-CsvPath` appears ignored | A nonempty `$OneDriveSiteList` takes precedence. Clear the list for report mode. |
| CSV candidate is skipped | Review archive status, unlicensed reason, and blocker allowlists. Unknown values are deliberately not treated as eligible. |
| Conflicting duplicate CSV rows | Obtain a fresh report or resolve the conflicting eligibility data. The script will not select an unblocked row over a blocked one. |
| URL does not match an admin host | Use the personal-site root and add the correct geo admin URL. Check for duplicate/misconfigured admin entries. |
| Vanity admin URL requires a mapping | Add its exact OneDrive hostnames to `$SPOAdminUrlMappings`. Do not assume a `-admin` to `-my` hostname replacement applies to a custom domain. |
| Token acquisition reports an unknown resource | Verify `TokenResourceUrl` with the tenant administrator. The configured vanity URL may not be a registered SharePoint OAuth resource. |
| Vanity endpoint works in a browser but fails in PnP | Verify that it supports authenticated SharePoint API/CSOM access, not just interactive redirects. Use the API-capable admin endpoint if necessary. |
| Audit write failed | Stop retrying deletions until disk space and access are fixed. Preserve the partial journal and reconcile live state for unfinished checkpoints. |
| Audit has more rows than the summary | Expected: mutation checkpoints and final results share the append-only journal. Use the last row per site. |
| A retry fails because the site no longer exists | Check Deleted Sites and the previous audit. The script does not automatically reconcile already-deleted targets. |

## Limitations and operational guidance

- The CSV is a snapshot. The script does not independently query Entra licensing or verify every retention/hold condition.
- Explicit-list mode puts candidate-selection responsibility on the administrator.
- `Remove-PnPTenantSite -Force` suppresses confirmation; this script does not request a second interactive approval.
- The script does not permanently purge sites or guarantee indefinite recoverability. Verify current service recovery policies and recovery procedures before deletion.
- A successful deletion command is not followed by a deleted-site verification query.
- Unlocking and deletion are not atomic, and failed deletion does not restore the previous lock state.
- Runs are sequential, but there is no cross-run locking. Avoid overlapping runs targeting the same sites.
- There is no persistent resume/reconciliation mechanism. Review prior journals and live state before retrying.
- Audit records have no individual checkpoint timestamp. Row order supplies event order within a run; the filename identifies the run time and unique ID.
- Automatic discovery processes one compatible report, not a tenant-wide inventory across all files.
- The report-helper retry settings do not configure a custom retry loop around PnP mutations.
- PnP.PowerShell is community-supported; using documented PnP commands is not equivalent to Microsoft product support for this script.

### Validation performed

Offline mocked checks covered duplicate/conflicting CSV rows, audit preflight and mid-run failure handling, unique filenames, discovery fallback, eligibility filters, geo routing, dry-run behavior, retention failures, continued processing, and failed unlock verification. PowerShell parsing and editor diagnostics were checked.

A read-only PnP integration check was also performed. These checks do not certify successful live PnP deletion/unlocking across tenants or prove that a target is free of retention restrictions.

### Suggested operating checklist

1. Obtain approval and confirm applicable retention/hold requirements.
2. Use certificate authentication and protect the app's broad permission.
3. Select a current report or independently verified explicit list.
4. Confirm every target's geo and root URL.
5. Keep `$PerformDeletes = $false` for the first run.
6. Review latest audit outcomes and skip/failure reasons.
7. Enable deletion only for the approved scope.
8. Review final outcomes and reconcile unfinished checkpoints.
9. Inspect any sites unlocked before a failed deletion.
10. Reset `$PerformDeletes = $false` and preserve the audit in an access-controlled location.

## Publishing checklist

Before uploading [the script](./Remove-UnlicensedArchivedOneDriveSites.ps1) and this README to GitHub:

- Replace environment-specific tenant IDs, application IDs, admin URLs, vanity-domain mappings, token resource URLs, certificate thumbprints, and sample personal-site URLs with placeholders.
- Ensure `$PerformDeletes = $false` and `$OneDriveSiteList = @()` in the published template.
- Do not include secrets, tokens, private keys, exported certificates with private keys, authentication traces, or sensitive logs.
- Exclude real report CSVs and audit files unless sanitized and approved for disclosure; they can expose personal-site URLs and account information.
- Publish only the intended files, not an entire customer working folder.
- Choose an appropriate repository license separately. This README does not assign a license or imply official Microsoft support.

The examples in this README use generic placeholders. Review the script separately before publication.

## References

- [PnP.PowerShell documentation](https://pnp.github.io/powershell/)
- [Get-PnPTenantSite](https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html)
- [Set-PnPTenantSite](https://pnp.github.io/powershell/cmdlets/Set-PnPTenantSite.html)
- [Remove-PnPTenantSite](https://pnp.github.io/powershell/cmdlets/Remove-PnPTenantSite.html)
- [Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html)
- [PnP authentication guidance](https://pnp.github.io/powershell/articles/authentication.html)
