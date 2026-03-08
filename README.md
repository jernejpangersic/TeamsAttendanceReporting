# Teams Attendance & Usage Reporting

PowerShell-based solution to extract Microsoft Teams meeting attendance data for education tenants via the Microsoft Graph API. Produces per-student, per-class attendance records with join/leave times and exports to Excel.

Built for M365 Education tenants (A3/A5) using School Data Sync, targeting scenarios where centralised, tenant-level attendance reporting is needed — such as fully remote learning.

## Features

- Extracts per-meeting attendance reports for all teachers in the tenant
- Computes attendance status: **Present**, **Late**, **Partial** (configurable thresholds)
- Exports to Excel (.xlsx) with one file per day
- Parallel processing with adaptive throttling (auto-adjusts delay based on 429 rate)
- Built-in Graph API retry handling (429/503/504) with exponential back-off
- Uses Call Records API as discovery layer with time-sharded parallel processing (`Get-AttendanceViaCallRecords-v5.ps1`, recommended for large tenants)
- Includes `Get-AttendanceViaCallRecords-v4.ps1` as a simpler alternative with sequential discovery

## Project Structure

```
├── config.sample.json               ← Template for tenant configuration
├── teachers.sample.json             ← Template showing teacher list format
├── recipients.sample.json           ← Template for department → email mapping
├── email-template.html              ← Customisable HTML email body template
├── Register-App.ps1                 ← Azure AD app registration helper
├── Sync-Teachers.ps1                ← Syncs teachers from Graph → teachers.json + recipients.json
├── Get-AttendanceViaCallRecords-v4.ps1 ← Call Records with parallel resolution + attendance
├── Get-AttendanceViaCallRecords-v5.ps1 ← Call Records with time-sharding + adaptive throttle (recommended)
├── archive/                          ← Older script versions (< v4, gitignored)
├── Split-AttendanceByDepartment.ps1 ← Splits Excel output into per-department files (logged via Write-Log)
├── Send-AttendanceReports.ps1       ← Emails per-department files to school IT admins
├── Test-Setup.ps1                   ← Environment validation script
├── output/                          ← Generated Excel files (gitignored)
└── logs/                            ← Execution logs (gitignored)
```

## Prerequisites

| Requirement | Details |
|---|---|
| **PowerShell 7+** | Required for parallel processing. [Install](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows) or run `winget install Microsoft.PowerShell` |
| **Microsoft.Graph module** | PowerShell SDK for Microsoft Graph |
| **ImportExcel module** | Excel export without requiring Office |
| **MicrosoftTeams module** | Needed once to create the Application Access Policy |
| **Azure AD app registration** | With application-level permissions (see below) |
| **Application Access Policy** | Teams policy granting the app access to meeting data |
| **Tenant admin consent** | Admin must consent to all app permissions |

## Setup — Step by Step

### 1. Install PowerShell 7+

The parallel scripts (`v3`–`v5`) require PowerShell 7 or later. Check your version:

```powershell
$PSVersionTable.PSVersion
```

If you see `5.x`, install PowerShell 7:

```powershell
winget install Microsoft.PowerShell
```

After installation, open a new terminal and run `pwsh` to start PowerShell 7.

### 2. Set Execution Policy

Scripts downloaded from the internet are blocked by default. Choose one option:

```powershell
# Option A — Persistent (recommended): allow local/remote-signed scripts
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
Get-ChildItem *.ps1 | Unblock-File

# Option B — Session only (nothing persisted):
Set-ExecutionPolicy Bypass -Scope Process
```

### 3. Install PowerShell Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
```

> **Note:** `MicrosoftTeams` is only needed once to create the Application Access Policy (Step 5). It is not required at runtime.

### 4. Register the Azure AD App

You can register the app automatically using the included helper script, or manually in the Azure portal.

#### Option A — Automated (as a tenant admin)

```powershell
.\Register-App.ps1
```

This will:
1. Connect to Microsoft Graph interactively (requires the `Application.ReadWrite.All` delegated scope)
2. Create an app registration named `TeamsAttendanceReporting`
3. Add all required permissions
4. Generate a 1-year client secret
5. Print the `tenantId`, `clientId`, and `clientSecret` values you'll need for `config.json`

After the script finishes, **grant admin consent** in the Azure portal:
> Entra ID → App registrations → TeamsAttendanceReporting → API permissions → Grant admin consent

#### Option B — Manual registration

1. Go to **Entra ID** → **App registrations** → **New registration**
2. Name: `TeamsAttendanceReporting`, Supported account types: Single tenant
3. Under **API permissions**, add the following **Application** permissions for Microsoft Graph:

| Permission | Purpose |
|---|---|
| `CallRecords.Read.All` | Discover call records tenant-wide |
| `OnlineMeetings.Read.All` | Resolve meeting objects by join URL |
| `OnlineMeetingArtifact.Read.All` | Read attendance reports |
| `User.Read.All` | Resolve user identities |
| `Group.Read.All` | Read teacher group membership |
| `Mail.Send` | Send per-department reports via email |

4. Click **Grant admin consent**
5. Under **Certificates & secrets**, create a new client secret and copy the **Value**

### 5. Create the Application Access Policy

The `OnlineMeetings.Read.All` and `OnlineMeetingArtifact.Read.All` permissions require a Teams Application Access Policy. This is a **one-time** setup step that requires a Teams Administrator.

```powershell
Connect-MicrosoftTeams

New-CsApplicationAccessPolicy -Identity "AttendanceReportingPolicy" `
    -AppIds "<your-client-id>" `
    -Description "Allow TeamsAttendanceReporting to read meetings and attendance"

Grant-CsApplicationAccessPolicy -PolicyName "AttendanceReportingPolicy" -Global
```

> **Important:** Wait **30 minutes** for the policy to propagate before running any attendance scripts. Running scripts before propagation completes will result in 403 errors.

### 6. Create config.json

```powershell
Copy-Item config.sample.json config.json
```

Edit `config.json` with your values:

```json
{
  "tenantId": "<your-azure-ad-tenant-id>",
  "clientId": "<your-app-registration-client-id>",
  "clientSecret": "<your-client-secret>",
  "teacherGroupId": "<security-group-id-containing-teachers>",
  "outputDir": "./output",
  "logsDir": "./logs",
  "retentionDays": 30,
  "timezone": "Asia/Riyadh",
  "lateThresholdMinutes": 10,
  "partialThresholdPercent": 50
}
```

| Field | Description |
|---|---|
| `tenantId` | Your Azure AD / Entra ID tenant GUID |
| `clientId` | The app registration's Application (client) ID |
| `clientSecret` | The client secret **value** (not the secret ID) |
| `teacherGroupId` | Object ID of the security group containing your teachers |
| `outputDir` | Where Excel files are saved (default: `./output`) |
| `logsDir` | Where log files are written (default: `./logs`) |
| `retentionDays` | Auto-delete output files older than this many days |
| `senderEmail` | Email address or shared mailbox to send reports from (requires `Mail.Send`) |
| `timezone` | IANA timezone for determining "yesterday" (e.g., `Asia/Riyadh`, `America/New_York`) |
| `lateThresholdMinutes` | Minutes after meeting start to mark a student as "Late" |
| `partialThresholdPercent` | % of meeting duration below which a student is marked "Partial" |

### 7. Sync the Teacher List

```powershell
.\Sync-Teachers.ps1
```

This reads the security group specified in `config.json` and writes `teachers.json` with each teacher's ID, name, email, and department.

### 8. Validate Your Setup

Run the included validation script to check that everything is correctly configured:

```powershell
.\Test-Setup.ps1
```

This checks PowerShell version, installed modules, config.json fields, teachers.json, and Graph API connectivity.

## Usage

```powershell
# Extract attendance for yesterday (default)
.\Get-AttendanceViaCallRecords-v5.ps1

# Extract attendance for a specific date
.\Get-AttendanceViaCallRecords-v5.ps1 "2026-03-01"

# Tune parallelism
.\Get-AttendanceViaCallRecords-v5.ps1 -TargetDate "2026-03-01" -ThrottleLimit 20 -DiscoveryShards 8
```

Output is saved to `output/callrecords_v5_YYYY-MM-DD.xlsx`.

### Split Reports by Department

```powershell
# Split a specific file into per-department Excel files
.\Split-AttendanceByDepartment.ps1 -ExcelPath .\output\callrecords_v5_2026-03-02.xlsx

# Split all Excel files in .\output
.\Split-AttendanceByDepartment.ps1

# Use a different config file
.\Split-AttendanceByDepartment.ps1 -ConfigPath .\config.test.json
```

Creates a date-based subfolder (e.g., `output/2026-03-02/`) with one Excel file per department. When multiple source files share the same date, each gets its own folder named after the full filename stem (e.g., `output/callrecords_v5_2026-03-02/`).

### Email Reports to School Admins

Before sending emails, ensure:
1. `senderEmail` is set in `config.json`
2. `recipients.json` exists with email addresses filled in (auto-generated by `Sync-Teachers.ps1`)
3. The `Mail.Send` permission is granted and admin-consented on the app registration

```powershell
# Send reports from the latest split output
.\Send-AttendanceReports.ps1

# Send from a specific folder
.\Send-AttendanceReports.ps1 -ReportDir .\output\callrecords_v5_2026-03-02

# Preview without sending
.\Send-AttendanceReports.ps1 -WhatIf

# Custom subject and template
.\Send-AttendanceReports.ps1 -Subject "Weekly Report - {{Date}}" -TemplatePath .\custom-email.html
```

The email body is loaded from `email-template.html`, which supports these placeholders:

| Placeholder | Replaced With |
|---|---|
| `{{Department}}` | Department name |
| `{{Date}}` | Report date |
| `{{RowCount}}` | Number of rows in the attached file |

Edit `email-template.html` freely to customise the email content.

### Script Parameters

| Parameter | Default | Description |
|---|---|---|
| `-TargetDate` | Yesterday | The date to extract attendance for |
| `-ConfigPath` | `.\config.json` | Path to the configuration file |
| `-ThrottleLimit` | 15 | Max concurrent parallel workers |
| `-DiscoveryShards` | 6 | Number of time-shards for Phase 1 discovery |

## Excel Output Columns

| Column | Example |
|---|---|
| Date | 2026-03-01 |
| TeacherName | Ms. Al-Rashid |
| TeacherEmail | a.rashid@school.edu |
| Department | Contoso High School |
| MeetingSubject | Biology - Grade 10 |
| MeetingStart | 2026-03-01 10:00:00 |
| MeetingEnd | 2026-03-01 10:45:00 |
| AttendeeName | Omar Hassan |
| AttendeeEmail | o.hassan@school.edu |
| JoinTime | 2026-03-01 10:02:00 |
| LeaveTime | 2026-03-01 10:44:00 |
| DurationMinutes | 42 |
| AttendanceStatus | Present |

## Attendance Status Logic

| Condition | Status |
|---|---|
| Join time > meeting start + threshold | **Late** (default: 10 min) |
| Duration < threshold % of meeting | **Partial** (default: 50%) |
| Otherwise | **Present** |

Thresholds are configurable in `config.json` (`lateThresholdMinutes`, `partialThresholdPercent`).

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `This script requires PowerShell 7+` | Running in Windows PowerShell 5.1 | Open `pwsh` instead of `powershell` |
| `403 Forbidden` on meeting/attendance calls | Application Access Policy not set or not yet propagated | Run the policy commands (Step 5) and wait 30 min |
| `401 Unauthorized` | Bad credentials or expired client secret | Check `clientId`/`clientSecret` in config.json |
| `Config file not found` | Missing `-ConfigPath` or wrong path | Ensure config.json exists and pass the correct path |
| Lots of `No meeting found (404)` | Normal — ad-hoc/P2P calls have no online meeting object | These are expected and logged as warnings |
| `Cannot convert null to type System.DateTime` | Older script version with unsafe DateTime casts | Use v5 which uses `TryParse` |
