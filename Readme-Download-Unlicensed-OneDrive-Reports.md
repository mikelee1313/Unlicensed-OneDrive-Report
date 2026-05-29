# Download Unlicensed OneDrive Reports (SharePoint Admin API)

Downloads the Unlicensed OneDrive Accounts report from SharePoint Online Admin Center using the same backend endpoint used by the UI Download report action.

Supports single-geo and multi-geo tenants by running one export per admin URL and saving one CSV per geo.

## Why This Script

The SharePoint Admin Center exposes this report in the portal, but many teams need:

- repeatable scheduled exports,
- app-only authentication (no interactive sign-in),
- multi-geo coverage in one run,
- retry handling for transient and throttling failures.

This script provides all of the above with no external PowerShell modules.

## What It Does

For each SharePoint admin URL in configuration, the script:

1. Acquires a SharePoint Online app-only access token.
2. Gets a form digest from /_api/contextinfo.
3. Calls POST /_api/SPO.Tenant/ExportToCSV with the unlicensed OneDrive filter/view payload.
4. Polls until the generated CSV is available.
5. Downloads the CSV locally with tenant label in the filename.

## Requirements

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or later (PowerShell 7 also works) |
| Dependencies | None (built-in cmdlets only) |
| Azure App Registration | Required for app-only auth |
| SharePoint API Permission | Application permission: Sites.FullControl.All (admin consent required) |
| Network Access | Access to login.microsoftonline.com and *.sharepoint.com |

## Authentication Modes

Set AuthType in the script configuration.

### Certificate (Recommended)

- Uses client assertion (JWT) signed with certificate private key.
- Certificate must exist in selected cert store and include private key access for the account running the script.

### Client Secret

- Uses standard client credentials flow with client secret.
- Simpler setup but less secure than certificate auth.

## Security Notes

- Do not commit real tenant IDs, client IDs, thumbprints, or secrets to a public repository.
- Before publishing, replace any live values in the configuration section with placeholders.
- Prefer certificate authentication over client secrets where possible.
- Consider rotating credentials regularly and using least privilege if your org supports narrowed scope alternatives.

## Configuration

Edit the configuration section at the top of the script.

| Setting | Description |
|---|---|
| tenantId | Microsoft Entra tenant ID |
| clientId | App registration client ID |
| AuthType | Certificate or ClientSecret |
| Thumbprint | Certificate thumbprint (Certificate mode) |
| CertStore | LocalMachine or CurrentUser |
| clientSecret | Client secret value (ClientSecret mode) |
| SPOAdminUrls | Array of SharePoint admin URLs (one per geo) |
| OutputFolder | Local folder for CSV output |
| MaxRetries | Retry limit for transient/throttled calls |
| InitialBackoffSec | Starting backoff delay |
| RequestTimeoutSec | Request timeout per call |
| SPOExportPollIntervalSec | Poll interval while waiting for CSV |
| SPOExportMaxWaitSec | Max wait time for CSV readiness |

## Multi-Geo Example

```powershell
$SPOAdminUrls = @(
  'https://contoso-admin.sharepoint.com'
  'https://contoso-eur-admin.sharepoint.com'
  'https://contoso-apc-admin.sharepoint.com'
)
```

Each admin URL gets its own SPO-scoped token and produces its own CSV.

## Usage

### 1) Configure values

Open the script and update tenant/app/auth/output settings.

### 2) Run

```powershell
pwsh .\Download-Unlicensed-OneDrive-Reports.ps1
```

or in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\Download-Unlicensed-OneDrive-Reports.ps1
```

### 3) Review output

The script prints progress for each admin URL and reports saved file paths at the end.

## Output

- One CSV per admin URL.
- File naming pattern:

```text
UnlicensedOneDrive_<tenantLabel>_Sites_<timestamp>.csv
```

Example:

```text
UnlicensedOneDrive_contoso_Sites_20260505164854854.csv
```

Typical report columns include:

- Display name
- Username
- Storage used (GB)
- Unlicensed due to
- Unlicensed on
- Deletion blocked by
- Owner email
- Deletion scheduled on
- Archive status
- Account provisioned for (UPN)
- URL

## Retry and Throttling Behavior

The script retries on:

- 429 (throttling)
- 502, 503, 504
- timeout / connection-closed transient network errors

Behavior:

- Honors Retry-After header for 429 when provided.
- Uses exponential backoff otherwise.
- Stops after MaxRetries and surfaces the error.

## Troubleshooting

### Access denied / auth failures

- Verify app permission Sites.FullControl.All is granted and admin-consented in SharePoint API.
- Confirm tenantId and clientId match the correct Entra tenant/app.
- For certificate mode, ensure certificate exists in selected store and private key is accessible.

### ExportToCSV returns no path

- Usually indicates an API payload or permission issue.
- Capture verbose output and validate the app can access the target admin URL.

### File not ready before timeout

- Increase SPOExportMaxWaitSec for large tenants.
- Keep SPOExportPollIntervalSec modest (for example 5 to 10 seconds).

### Multi-geo missing output

- Ensure each geo admin URL is valid and reachable.
- Confirm the app has effective permissions for each geo endpoint.

## Known Limitations

- The script uses an internal backend endpoint format used by admin center behavior. Microsoft may change this API behavior without notice.
- Configuration is inline in the script (not parameterized via command-line arguments).

## Recommended Enhancements

- Convert configuration to formal script parameters.
- Add optional logging to file (JSON and/or transcript).
- Add optional email/Teams notification on completion.
- Add exit codes for CI/CD or scheduler health checks.
- Add Pester tests for helper functions.

## Publishing Checklist (GitHub)

Before posting publicly:

1. Replace all real IDs, thumbprints, URLs, and secrets with placeholders.
2. Verify no sensitive values exist in commit history.
3. Add a .gitignore entry for exported CSV files if needed.
4. Include licensing for your organization policy (MIT shown below as example).

## License

MIT
