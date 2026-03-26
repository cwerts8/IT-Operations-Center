# IT Operations Center

A comprehensive PowerShell WPF GUI tool for managing Exchange Online, Active Directory, Microsoft 365, and Intune operations. Built for the Desktop Engineering team as a centralized platform for IT administrators to perform complex user management, account provisioning, and system operations through a single interface — no command line required.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Modules](#modules)
  - [Exchange Online](#exchange-online)
  - [Microsoft 365](#microsoft-365)
  - [Active Directory](#active-directory)
  - [Device Management](#device-management)
- [Architecture Notes](#architecture-notes)
- [Version History](#version-history)

---

## Overview

IT Operations Center launches as a WPF GUI window. Exchange Online connection is optional — the tool opens immediately and prompts you to connect only when you access an Exchange-dependent module. The `ExchangeOnlineManagement` module is loaded on-demand at connection time to avoid assembly conflicts with Microsoft Graph modules used by Intune and other features. AD and Intune modules connect independently using their own module requirements.

All modules share a consistent interface pattern: search or input at the top, results in a sortable/filterable DataGrid, action buttons below, and a live status bar at the bottom. Most modules support Excel export.

---

## Prerequisites

| Requirement | Used By | Install Method |
|---|---|---|
| PowerShell 5.1+ | All modules | Built into Windows |
| ExchangeOnlineManagement module | Exchange Online modules | `Install-Module ExchangeOnlineManagement` |
| Active Directory RSAT module | All AD modules | Settings → Apps → Optional Features → RSAT: Active Directory Domain Services |
| Microsoft.Graph.Authentication | TAP generation, Intune | `Install-Module Microsoft.Graph.Authentication` |
| Microsoft.Graph.Users | TAP generation | `Install-Module Microsoft.Graph.Users` |
| Microsoft.Graph.DeviceManagement | Intune module | `Install-Module Microsoft.Graph.DeviceManagement` |
| ImportExcel module | All Excel exports | Auto-installed on first run if missing |

> **Note:** The script checks for `ImportExcel` on startup and installs it automatically if not present. All other modules must be installed manually before using their respective features.

> **Module Load Order:** `ExchangeOnlineManagement` is intentionally loaded on-demand when you connect to Exchange Online — not at startup. This prevents `Azure.Core.dll` version conflicts that would otherwise block Microsoft Graph modules (Intune, TAP) from loading in the same session.

---

## Getting Started

1. Clone or download `IT-Operations-Center.ps1` to your machine.
2. Right-click the file and choose **Run with PowerShell**, or run it from a PowerShell terminal:

```powershell
.\IT-Operations-Center.ps1
```

3. The GUI will launch immediately. The Exchange Online connection status indicator in the header shows **red (Disconnected)** by default.
4. Click **Connect** in the header bar to authenticate to Exchange Online via browser when you need Exchange features.
5. AD and Device Management modules connect independently — no EXO connection required for those.

> **Execution Policy:** If you receive a script execution error, run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` in PowerShell first.

---

## Modules

### Exchange Online

All Exchange Online modules require an active EXO connection. The tool will prompt you to connect if you attempt to open one while disconnected.

#### Mailbox Permissions
Manage Full Access and Send As permissions for any mailbox.
- Load current permissions for any mailbox (by email or username)
- Add Full Access permissions (AutoMapping disabled by default)
- Add Send As permissions
- Remove individual permissions
- GUID-to-name resolution for AD group entries
- Double-click any entry to view AD properties
- Export permissions to Excel

#### Calendar Permissions
Manage calendar delegation for any mailbox.
- 7 permission levels: AvailabilityOnly, LimitedDetails, Reviewer, Author, NonEditingAuthor, Editor, Owner
- Add, modify, and remove calendar access
- Export to Excel

#### Automatic Replies (Out of Office)
Configure Out of Office messages without needing to know HTML.
- Rich text editor with Bold, Italic, Underline formatting toolbar
- Automatic HTML conversion from formatted text
- Internal and external reply messages (separate content for each)
- Three states: Disabled, Enabled (always on), Scheduled (date/time range)
- External audience controls: All senders or Contacts only
- Color-coded status display (Gray = Disabled, Green = Enabled, Orange = Scheduled)

#### Message Trace / Tracking
Troubleshoot email delivery issues.
- Search by sender, recipient, subject, message ID, or delivery status
- Date range options: Last 24 hours, Last 7 days, or custom range (up to 10 days)
- Status filters: All, Delivered, Failed, Pending, Quarantined, FilteredAsSpam
- Configurable result limits: 100, 1,000, or 5,000 messages
- Detailed delivery event timeline per message
- Copy Message ID to clipboard for deeper investigation
- Export results to Excel

> **Distribution Group Note:** Message Trace searches show emails sent *to* the group address. External emails delivered to individual group members may not appear in results.

---

### Microsoft 365

#### Generate Temporary Access Pass (TAP)
Generate time-limited passcodes for passwordless authentication onboarding (Windows Hello, Microsoft Authenticator, FIDO2 keys).
- Lookup user by SAMAccountName or UPN
- Configurable lifetime: 10 to 43,200 minutes (up to 30 days)
- One-time use or multi-use options
- Connects to Microsoft Graph automatically with `UserAuthenticationMethod.ReadWrite.All` scope
- TAP displayed once with one-click copy to clipboard
- Shows user display name, expiry time, and usage type after generation

---

### Active Directory

All AD modules require the Active Directory RSAT PowerShell module.

#### AD Group Members Viewer
View all members of any AD group (Security Groups, Distribution Groups, etc.).
- Search by group name or email address
- Displays members with enriched details: Display Name, Email, Title, Department, SAMAccountName, Object Type
- Supports all member types: Users, Groups, Computers, Contacts
- Shows Group Category (Security/Distribution) and Scope (DomainLocal/Global/Universal)
- Copy all member emails to clipboard in Outlook-ready semicolon format
- Export to Excel
- Double-click any member to open their AD Properties

#### User Group Memberships
View all AD groups a specific user belongs to.
- Lookup by SAMAccountName or UPN
- Displays Group Name, Category, Scope, and Description for each group
- User info banner shows Display Name, Department, Email, and total group count
- Results sorted alphabetically
- Export to Excel

#### User Group Comparison
Compare the group memberships of two AD users side by side — useful when onboarding a new team member or converting a role.
- Enter any two usernames or UPNs
- **Side-by-Side View tab:** Two independent scrollable grids (User A | User B) with a draggable splitter
- **Comparison View tab:** Color-coded unified list
  - Green — Shared by both users
  - Blue — Only User A has this group
  - Purple — Only User B has this group
- Summary bar shows shared count, unique-to-A count, unique-to-B count, and total groups per user

#### Employee Conversion
Convert an employee between role types (e.g., Consultant → Full Time Employee) using a template-based system.
- Uses `_Templatexxx` accounts to define group sets per role type
- Select a "from template" (current role) and "to template" (new role)
- Surgical group management: only removes groups that came from the source template
- Preserves any custom groups not associated with either template
- Adds all groups from the destination template
- Automatically moves the user to the new template's OU
- Preview panel shows exactly which groups will be added, removed, and preserved before applying
- Comprehensive logging of all changes

#### Locked Out Users
Manage locked out AD accounts with bulk unlock support.
- Loads all currently locked out accounts from Active Directory
- Real-time search and filtering across Name, Username, Email, Department, Title
- Multi-select rows for bulk unlock operations
- Displays lockout time, bad logon count, and user details per account
- Automatic refresh after unlock
- Export locked out user list to Excel

#### Export Active Users Report
One-click export of all enabled AD user accounts.
- Pulls from Active and Consultants OUs
- Automatically filters out test accounts (`test` prefix or `t-` prefix)
- Exports: Name, SamAccountName, Email, Department, Title, Office (desk location)
- Sorted by Name ascending
- User-selectable save location
- Excel output with auto-formatting, column filters, and frozen header row

#### AD Properties Viewer
Available from any module — double-click any user or group entry to open a read-only ADUC-style properties window.
- **General tab:** Display Name, Email, Title, Department, Office, Company
- **Contact Information tab:** Phone, Mobile, Address
- **Organization tab:** Manager, Direct Reports, Group Memberships
- **Account tab:** UPN, SAMAccountName, Distinguished Name, GUID, Created/Modified dates
- Copy email address to clipboard

---

### Device Management

#### Intune Managed Devices
View and manage all devices enrolled in Microsoft Intune.
- Connects to Microsoft Graph automatically
- Dashboard with real-time counts by OS type (iOS, Android, Windows, macOS)
- Filter by device type using checkboxes (applied automatically on load)
- Real-time search box: filter by device name, user, model, serial number, IMEI, or OS without new API calls
- Comprehensive device details: Name, User, IMEI, Serial, Model, OS Version, Storage, Enrollment Date, Last Sync, Compliance State
- Compliance status color coding in Excel export (Compliant = green, Non-Compliant = red, Jailbroken = orange)
- IMEI/MEID exported as full numbers (no scientific notation)

#### IP Network Scanner
Scan IP ranges to discover active devices on the network.
- Specify a start and end IP address for any range
- Parallel scanning with 50 concurrent threads (~50 IPs/sec vs ~1 IP/sec sequential)
- Tests connectivity using `Test-Connection` with 1-second timeout
- Automatic hostname resolution for online devices
- MAC address detection via ARP table lookup
- Response time measurement in milliseconds
- Real-time progress with online/offline counters
- Color-coded results: Green = Online, Red = Offline
- Prompts for confirmation on ranges over 1,000 addresses
- Export results to Excel

---

## Architecture Notes

**Single-file application.** The entire tool is one `.ps1` file. XAML for each module window is defined inline as a here-string and parsed at runtime with `XamlReader`.

**`$syncHash` pattern.** All GUI controls are stored in a shared `$syncHash` hashtable to ensure they remain accessible across event handler scopes. New modules follow the same pattern: register controls in `$syncHash` after the XAML block, then reference them inside `Add_Click` handlers.

**On-demand module loading.** `ExchangeOnlineManagement` is loaded only when the user explicitly connects to Exchange Online — not at startup. This is intentional: loading it at startup caused an `Azure.Core.dll` version conflict that prevented Microsoft Graph modules (used by Intune and TAP generation) from loading in the same session.

**Array handling.** PowerShell automatically unwraps single-item pipeline results from a collection to a bare object. All `Get-AD*` calls and `Sort-Object` pipelines that feed DataGrid `ItemsSource` are wrapped in `@()` and cast to `[System.Collections.ObjectModel.ObservableCollection[object]]` to prevent WPF binding failures when results contain only one item.

**Excel exports** use the `ImportExcel` module (auto-installed on first run). Exports include `AutoSize`, `AutoFilter`, `FreezeTopRow`, and `BoldTopRow` by default.

**Logging.** All significant actions are written to the activity log panel in the main window via the `Write-Log` function. Timestamps are in `HH:mm:ss` format.

---

## Version History

| Version | Date | Summary |
|---|---|---|
| 1.0.0 | 10-06-25 | Initial release — Mailbox permissions, Calendar permissions, AD Group Members |
| 2.0 | 10-08-25 | Excel export, GUID resolution, AD Properties double-click viewer |
| 2.6 | 10-09-25 | Error handling improvements, Copy Emails button, progress indicators |
| 2.6.1 | 10-09-25 | Company logo and dynamic version text in header |
| 2.7.0 | 10-13-25 | Automatic Replies module with rich text editor |
| 2.8.0 | 10-13-25 | Optional EXO connection — GUI launches without requiring authentication |
| 2.8.1 | 10-13-25 | Bug fix: single-delegate array handling in Mailbox and Calendar modules |
| 2.9.0 | 10-16-25 | AD Group Members switched from EXO to AD module — supports all group types |
| 3.2.0 | 10-23-25 | Message Trace / Tracking module |
| 3.3.0 | 10-27-25 | IP Network Scanner with parallel 50-thread scanning |
| 3.4.0 | 10-28-25 | Export Active Users Report |
| 3.5.0 | 10-28-25 | Intune Mobile Devices module with Microsoft Graph integration |
| 4.0.0 | 10-29-25 | Major UI redesign — consolidated header, reorganized Management Options into 4 categories |
| 4.1.0 | 10-29-25 | Intune module enhanced — all device types, real-time search, dynamic filtering |
| 4.2.0 | 11-10-25 | Generate Temporary Access Pass (TAP) — Microsoft 365 section added |
| 4.3.0 | 11-12-25 | Employee Conversion module |
| 4.4.0 | 11-14-25 | Locked Out Users module with bulk unlock |
| 4.5.0 | 11-14-25 | Documentation and version tracking improvements |
| 4.5.1 | 11-14-25 | Bug fix: AD Group Members viewer failed when group had only one member |
| 4.6.0 | 11-14-25 | User Group Memberships viewer and User Group Comparison tool |
| 4.7.1 | 03-17-26 | Fix: `ExchangeOnlineManagement` now loads on-demand to resolve `Azure.Core.dll` conflict with Microsoft Graph modules |

---

*Maintained by Craig Werts — Desktop Engineering*
