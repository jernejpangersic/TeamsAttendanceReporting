# Teams Attendance & Usage Reporting

PowerShell-based solution to extract Microsoft Teams meeting attendance data for education tenants via the Microsoft Graph API. Produces per-student, per-class attendance records with join/leave times and exports to Excel.

Built for M365 Education tenants (A3/A5) using School Data Sync, targeting scenarios where centralised, tenant-level attendance reporting is needed — such as fully remote learning.

## Features

- Extracts per-meeting attendance reports for all teachers in the tenant
- Computes attendance status: **Present**, **Late**, **Partial** (configurable thresholds)
- Exports to Excel (.xlsx) with one file per day
- Parallel processing of teachers (throttle-limited)
- Built-in Graph API retry handling (429/503/504)
- Two extraction strategies:
  - `Get-Attendance.ps1` — polls each teacher's online meetings (straightforward)
  - `Get-AttendanceViaCallRecords.ps1` — uses Call Records as discovery layer (optimised for large tenants with many inactive teachers)

## Project Structure

```
├── config.sample.json               ← Template for tenant configuration
├── teachers.sample.json             ← Template showing teacher list format
├── Register-App.ps1                 ← Azure AD app registration helper
├── Sync-Teachers.ps1                ← Syncs teachers from Graph → teachers.json
├── Get-Attendance.ps1               ← Main script: meetings → attendance → Excel
├── Get-AttendanceViaCallRecords.ps1 ← Alternate script using Call Records API
├── output/                          ← Generated Excel files (gitignored)
└── logs/                            ← Execution logs (gitignored)
```

## Prerequisites

1. **PowerShell 7+** (pwsh)
2. **Microsoft Graph PowerShell SDK**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```
3. **ImportExcel module** (for .xlsx export)
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```
4. **Azure AD App Registration** with the following **application** permissions (admin consent required):
   | Permission | Purpose |
   |---|---|
   | `OnlineMeetingArtifact.Read.All` | Read attendance reports |
   | `OnlineMeeting.Read.All` | List meetings per user |
   | `User.Read.All` | Resolve user identities |
   | `Group.Read.All` | Read teacher group membership |
   | `CallRecords.Read.All` | *(Only for CallRecords variant)* |

5. **Application access policy** — required for OnlineMeeting/Artifact permissions:
   ```powershell
   New-CsApplicationAccessPolicy -Identity "AttendanceReportingPolicy" `
       -AppIds "<your-client-id>" `
       -Description "Allow TeamsAttendanceReporting to read meetings and attendance"

   Grant-CsApplicationAccessPolicy -PolicyName "AttendanceReportingPolicy" -Global
   ```

## Setup

1. **Register the Azure AD app** (or run `Register-App.ps1` interactively as a tenant admin):
   ```powershell
   .\Register-App.ps1
   ```

2. **Create `config.json`** from the sample:
   ```powershell
   Copy-Item config.sample.json config.json
   ```
   Edit `config.json` with your tenant ID, client ID, client secret, and teacher group ID.

3. **Sync the teacher list**:
   ```powershell
   .\Sync-Teachers.ps1
   ```
   This creates `teachers.json` from the configured security group.

## Usage

```powershell
# Extract attendance for yesterday (default)
.\Get-Attendance.ps1

# Extract attendance for a specific date
.\Get-Attendance.ps1 -TargetDate "2026-03-01"

# Use the Call Records variant (better for large tenants)
.\Get-AttendanceViaCallRecords.ps1 -TargetDate "2026-03-01"
```

Output is saved to `output/attendance_YYYY-MM-DD.xlsx`.

## Excel Output Columns

| Column | Example |
|---|---|
| Date | 2026-03-01 |
| TeacherName | Ms. Al-Rashid |
| TeacherEmail | a.rashid@school.edu |
| MeetingSubject | Biology - Grade 10 |
| MeetingStart | 2026-03-01 10:00:00 |
| MeetingEnd | 2026-03-01 10:45:00 |
| StudentName | Omar Hassan |
| StudentEmail | o.hassan@school.edu |
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
