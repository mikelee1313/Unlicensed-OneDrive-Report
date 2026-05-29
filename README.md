# Unlicensed OneDrive Report

Comprehensive PowerShell toolkit to identify, track, and manage unlicensed OneDrive accounts across Microsoft 365 tenants using Microsoft Graph API—no SharePoint Online PowerShell module, no per-geo tokens, and no manual admin center navigation required.

---

## Overview

When users lose OneDrive/SharePoint licenses (removal, deletion, or plan changes), Microsoft enforces an archival timeline:
- **Day 60**: OneDrive becomes read-only
- **Day 93**: OneDrive is archived or deleted (depending on billing settings)

This repository contains three complementary PowerShell scripts that work together to help tenant admins identify unlicensed accounts, download official reports, and automate secure offboarding.

---

## Scripts at a Glance

### 1. **Get-UnlicensedOneDriveReport.ps1**
The comprehensive Microsoft Graph–based discovery and reporting engine.

**What it does:**
- Scans all active Entra ID users to find those without active OneDrive/SharePoint plans
- Discovers soft-deleted users in the Entra 30-day recycle bin
- Finds already-archived personal OneDrive sites (for users purged >30 days ago)
- Enriches findings with audit log license-removal dates
- Calculates exact milestone dates (Day 60 read-only, Day 93 archive)
- Assigns traffic-light urgency labels (CRITICAL, WARNING, MONITOR, OK, ARCHIVED)
- Sends optional HTML alert emails to admins before sites go read-only or are archived
- Estimates archive/reactivation costs per account

**Key features:**
- ✅ **Multi-geo aware** — single token covers all regions (NAM, APC, CAN, DEU, GBR, IND, JPN)
- ✅ **No SPO PowerShell module required** — pure Microsoft Graph API
- ✅ **Certificate or Client Secret auth** — both flows supported
- ✅ **Detects three populations** — active users + soft-deleted users + already-archived sites
- ✅ **Throttle handling** — exponential backoff with Retry-After support
- ✅ **Email alerts** — sends HTML notifications before critical milestones
- ✅ **UTF-8 BOM CSV output** — Excel-safe encoding

**Output:**
- Timestamped CSV: `UnlicensedOneDrive_<yyyyMMddHHmmss>.csv`
- Columns: UserSource, DisplayName, UPN, UnlicensedDate, DaysUntilReadOnly, DaysUntilArchive, UrgencyStatus, StorageUsedGB, DriveUrl, CostEstimates, and more
- Optional HTML alert emails sent to configured admin addresses/groups

**Documentation:**  
[README-Get-UnlicensedOneDriveReport.md](./README-Get-UnlicensedOneDriveReport.md)

---

### 2. **Download-Unlicensed-OneDrive-Reports.ps1**
A lightweight wrapper around the SharePoint Admin Center's native Unlicensed OneDrive report endpoint.

**What it does:**
- Downloads the built-in "Unlicensed OneDrive Accounts" report from SharePoint Admin Center
- Uses app-only authentication (no interactive sign-in required)
- Supports multi-geo tenants with automated per-geo CSV export
- Includes retry logic for throttling and transient failures
- Polls for CSV generation and downloads when ready

**Key features:**
- ✅ **Uses official SPO API endpoint** — same backend as the Admin Center UI
- ✅ **App-only auth** — no user interaction, ideal for automation
- ✅ **Multi-geo support** — one export per admin URL
- ✅ **Retry/throttle handling** — honors Retry-After headers
- ✅ **No PowerShell modules required** — built-in cmdlets only

**Output:**
- Per-geo CSV files: `UnlicensedOneDrive_<tenantLabel>_Sites_<timestamp>.csv`
- Columns: Display name, Username, Storage used (GB), Unlicensed due to, Unlicensed on, Owner email, Deletion scheduled on, Archive status, URL

**Use case:**  
When you want a lightweight, official report without the deep audit log and cost analysis features of Script #1.

**Documentation:**  
[Readme-Download-Unlicensed-OneDrive-Reports.md](./Readme-Download-Unlicensed-OneDrive-Reports.md)

---

### 3. **OneDriveOffboarding.ps1**
Automates the secure handoff of unlicensed/deleted user OneDrives to managers.

**What it does:**
- Accepts a CSV list of users to offboard
- Resolves each user and their manager in Microsoft Graph
- Grants manager Site Collection Admin (SCA) access to the OneDrive
- Optionally locks the site to read-only
- Inventories all active sharing links
- Sends an HTML manager notification email with site details and link metadata
- Exports a run summary CSV with success/failure tracking

**Key features:**
- ✅ **App-only Graph + SharePoint auth** — no delegated user permissions
- ✅ **Flexible auth** — Certificate-based client credentials
- ✅ **Sharing link inventory** — discovers links via SharePoint SharingLinks groups
- ✅ **Audit log fallback** — calculates offboarding baseline date from audit logs if CSV date missing
- ✅ **Manager email notifications** — rich HTML with site and link details
- ✅ **Run tracking** — success/failure CSV for auditing

**Input:**
- CSV with columns: `UserPrincipalName`, optional `DeletedDate`, optional `OneDriveUrl`

**Output:**
- Timestamped CSV: `OneDrive_Manager_Handoff_<yyyyMMdd_HHmmss>.csv`
- Columns: UserPrincipalName, Manager, OneDriveUrl, SiteLockState, SharingLinkCount, NotificationSent, Notes
- HTML manager emails sent via Graph sendMail

**Use case:**  
When offboarding deleted or unlicensed users, ensure managers can access and secure data while notifying them of sharing arrangements.

**Documentation:**  
[Readme-OneDriveOffboarding.md](./Readme-OneDriveOffboarding.md)

---

## Which Script Should I Use?

| Scenario | Script | Why |
|----------|--------|-----|
| **Full discovery + reporting + cost analysis** | Script #1 (Get-UnlicensedOneDriveReport) | Identifies all three populations (active/soft-deleted/archived), calculates milestones, estimates costs, sends alerts |
| **Quick official report, minimal overhead** | Script #2 (Download-Unlicensed-OneDrive-Reports) | Uses native SPO API, lightweight, multi-geo aware, no deep analysis |
| **Offboard users + transfer access to managers** | Script #3 (OneDriveOffboarding) | Grants SCA, locks sites, inventories sharing, notifies managers |
| **Monitor unlicensed OneDrives on a schedule** | Scripts #1 + #3 | Use #1 to discover, then feed results to #3 to automate offboarding workflows |

---

## Prerequisites

### PowerShell
- PowerShell 5.1 or later (Windows)
- PowerShell 7.x recommended for Script #3

### Azure App Registration
A single app registration in the home tenant with the following **Application** permissions (admin consent required):

| Permission | Script #1 | Script #2 | Script #3 | Purpose |
|---|---|---|---|---|
| `User.Read.All` | ✅ | — | ✅ | Enumerate users and licenses |
| `Directory.Read.All` | ✅ | — | ✅ | Soft-deleted users + manager references |
| `Files.Read.All` | ✅ | — | ✅ | OneDrive drive metadata |
| `AuditLog.Read.All` | ⚠️ Optional | — | ⚠️ Optional | License removal dates + audit events |
| `Sites.Read.All` | ⚠️ Optional | ✅ | — | Archived site enumeration + SCA operations |
| `Mail.Send` | ⚠️ Optional | — | ✅ | Send alert/manager emails |
| `Sites.FullControl.All` | — | ✅ | ✅ | SPO admin operations (grant SCA, set lock state) |

### Authentication
All scripts support:
- **Certificate-based** client credentials (recommended for production)
- **Client Secret** credentials

### For Script #3 Only
- **PnP.PowerShell** module (install via `Install-Module PnP.PowerShell`)

---

## Quick Start

### 1. Configure an App Registration
1. Create an app registration in Entra ID
2. Add the required **Application** permissions (see table above)
3. Grant admin consent
4. Upload certificate public key or create a client secret
5. Note: Tenant ID, Client ID, Certificate Thumbprint/Secret

### 2. Install/Run Script #1 (Full Discovery Report)
```powershell
# Edit the CONFIGURATION section at the top of the script
$tenantId  = 'your-tenant-id'
$clientId  = 'your-app-client-id'
$Thumbprint = 'your-cert-thumbprint'

# Then run:
.\Get-UnlicensedOneDriveReport.ps1
