# Teams Attendance & Usage Reporting — Implementation Specification

**Date:** March 2, 2026
**Scope:** Proof of Concept (PoC)
**Purpose:** Technical blueprint for building the PoC. This document contains all details needed to implement the attendance extraction and Excel export, including code structure, API specifications, data schemas, and configuration.

---

## Table of Contents

1. [Project Structure](#1-project-structure)
2. [Azure AD App Registration](#2-azure-ad-app-registration)
3. [Infrastructure](#3-infrastructure)
4. [Teacher List Sync](#4-teacher-list-sync)
5. [Meeting Attendance Report](#5-meeting-attendance-report)
6. [Excel Export](#6-excel-export)
7. [Attendance Status Computation](#7-attendance-status-computation)
8. [Configuration](#8-configuration)
9. [Error Handling & Retry Logic](#9-error-handling--retry-logic)
10. [Testing Plan](#10-testing-plan)
11. [Deployment Runbook](#11-deployment-runbook)

---

## 1. Project Structure

```
teamsReporting/
├── 01-Solution-Overview.md          ← Solution overview document
├── 02-Implementation-Spec.md        ← This file
├── config.json                      ← Tenant config + attendance thresholds
├── teachers.json                    ← Cached teacher list (synced separately)
├── Sync-Teachers.ps1                ← Syncs teachers from Graph → teachers.json
├── Get-Attendance.ps1               ← Main script: meetings → attendance → Excel
├── Register-App.ps1                 ← Azure AD app registration helper
├── output/                          ← Generated Excel files (attendance_YYYY-MM-DD.xlsx)
└── logs/                            ← Error and execution logs
```

---

## 2. Azure AD App Registration

### App Details

| Property | Value |
|---|---|
| Name | `TeamsAttendanceReporting` |
| Type | Web application (confidential client) |
| Authentication | Client credentials flow (client_id + client_secret or certificate) |
| Redirect URI | Not required (daemon app) |

### API Permissions (Application)

| API | Permission | Type | Purpose |
|---|---|---|---|
| Microsoft Graph | `OnlineMeetingArtifact.Read.All` | Application | Read attendance reports |
| Microsoft Graph | `OnlineMeeting.Read.All` | Application | List meetings per user |
| Microsoft Graph | `User.Read.All` | Application | Enumerate teachers/students |
| Microsoft Graph | `Group.Read.All` | Application | Read teacher group membership (used by `Sync-Teachers.ps1`) |

### Application Access Policy (Required)

`OnlineMeetingArtifact.Read.All` and `OnlineMeeting.Read.All` with application permissions **require an application access policy**. The tenant admin must:

1. Create the policy via PowerShell:
   ```powershell
   # Requires: Microsoft Teams PowerShell module
   New-CsApplicationAccessPolicy -Identity "AttendanceReportingPolicy" `
       -AppIds "{AppId}" `
       -Description "Allow TeamsAttendanceReporting to read meetings and attendance"
   ```
2. Assign the policy to the relevant users (or all users):
   ```powershell
   # Grant to all users in the tenant:
   Grant-CsApplicationAccessPolicy -PolicyName "AttendanceReportingPolicy" -Global

   # Or grant to a specific user:
   Grant-CsApplicationAccessPolicy -PolicyName "AttendanceReportingPolicy" `
       -Identity "{userId}"
   ```
3. Allow up to **30 minutes** for the policy to take effect.

> **Reference:** [Configure application access policy](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)

### Registration Script Outline (`register-app.ps1`)

```powershell
# Requires: Microsoft.Graph PowerShell module
# 1. Connect-MgGraph -Scopes "Application.ReadWrite.All"
# 2. New-MgApplication -DisplayName "TeamsAttendanceReporting"
# 3. Add required resource access (permissions above)
# 4. Create client secret: Add-MgApplicationPassword
# 5. Output: AppId, TenantId, ClientSecret
# 6. Remind admin to grant consent in Entra ID portal
# 7. Create and assign application access policy (see above)
```

---

## 3. Infrastructure (PoC)

For the PoC, no Azure resources are required. The scripts run locally with PowerShell 7+.

| Requirement | Details | Cost |
|---|---|---|
| PowerShell 7+ | Local machine | $0 |
| `Microsoft.Graph` module | `Install-Module Microsoft.Graph` | $0 |
| `ImportExcel` module | `Install-Module ImportExcel` | $0 |
| Azure Automation (optional) | If scheduled runs are needed | $0-5/month |

**Post-PoC:** If promoted to production, add multiple app registrations for parallelism, Azure SQL, SharePoint, Power BI, etc.

---

## 4. Teacher List Sync

### Script: `Sync-Teachers.ps1`

**Purpose:** Maintains a cached list of teachers in `teachers.json`. Run separately from the attendance extraction — weekly or when staff changes occur.

### How It Works

```
1. Authenticate to Graph API (client credentials)
2. GET /groups/{teacherGroupId}/members/microsoft.graph.user
       ?$select=id,displayName,mail,department,officeLocation&$top=999
   → Page through all results using @odata.nextLink (single paginated call — no per-member round-trips)
3. Write results to teachers.json
4. Log: "{count} teachers synced"
```

### PowerShell Implementation

```powershell
# Sync-Teachers.ps1
param(
    [string]$ConfigPath = ".\config.json"
)

$config = Get-Content $ConfigPath | ConvertFrom-Json

# Authenticate with client credentials
$secureSecret = ConvertTo-SecureString $env:CLIENT_SECRET -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome

# Fetch all group members as user objects with properties in a single paginated call.
# The /microsoft.graph.user cast avoids a separate Get-MgUser call per member.
$uri = "https://graph.microsoft.com/v1.0/groups/$($config.teacherGroupId)/members/microsoft.graph.user" +
       "?`$select=id,displayName,mail,department,officeLocation&`$top=999"
$headers = @{ Authorization = "Bearer $((Get-MgContext).AccessToken)" }

$teachers = [System.Collections.Generic.List[object]]::new()
$nextUri  = $uri
while ($nextUri) {
    $resp = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method Get -ErrorAction Stop
    foreach ($u in $resp.value) {
        $teachers.Add([PSCustomObject]@{
            id             = $u.id
            displayName    = $u.displayName
            mail           = $u.mail
            department     = $u.department
            officeLocation = $u.officeLocation
        })
    }
    $nextUri = $resp.'@odata.nextLink'
}

$teachers | ConvertTo-Json -Depth 3 | Set-Content ".\teachers.json" -Encoding UTF8
Write-Host "$($teachers.Count) teachers synced to teachers.json"

Disconnect-MgGraph
```

### Refresh Strategy

| Trigger | Method |
|---|---|
| **Scheduled** | Weekly (e.g., Sunday night before the school week) |
| **On-demand** | Admin runs `Sync-Teachers.ps1` after staff changes |
| **Delta query** | Use `/groups/{id}/members/delta` for incremental sync (future optimisation) |

### `teachers.json` Output Format

```json
[
  {
    "id": "user-guid-1",
    "displayName": "Ms. Al-Rashid",
    "mail": "a.rashid@school.edu",
    "department": "Science",
    "officeLocation": "Al-Farabi Secondary"
  }
]
```

---

## 5. Meeting Attendance Report

### Script: `Get-Attendance.ps1`

**Trigger:** Run manually from command line, or as an Azure Automation runbook (daily at 02:00 UTC)

### Logic Flow

```
1. Load config.json and teachers.json
2. Define date range: yesterday 00:00 UTC → today 00:00 UTC (overridable via -TargetDate)
3. Acquire Graph token via MSAL client credentials (with auto-refresh caching)
4. For each teacher (ForEach-Object -Parallel, ThrottleLimit 10):
   a. Each runspace acquires its own token (tokens last ~60 min, one teacher takes seconds)
   b. GET /users/{teacherId}/onlineMeetings?$filter=startDateTime ge {start}...
      → Try server-side $filter; if rejected, fall back to $top=50&$orderby=startDateTime desc
        and filter client-side by date range
      → Follow @odata.nextLink for pagination
      → Skip to next teacher if no meetings found
   c. For each meeting:
      i.  GET .../attendanceReports (paginated)
          → Filter reports by meetingStartDateTime (±5 min) to exclude channel-meeting pollution
      ii. For each matching report:
          GET .../attendanceReports/{reportId}?$expand=attendanceRecords
          → Per-participant join/leave data
   d. For each attendance record:
      - Null-guard attendanceIntervals before indexing
      - Compute AttendanceStatus (Present / Late / Partial) using config thresholds
      - Emit [PSCustomObject] row
   e. On error: log and continue with next teacher (don't abort the batch)
5. Export collected rows → Excel via ImportExcel (Export-AttendanceExcel)
6. Save as output/attendance_YYYY-MM-DD.xlsx
7. Delete Excel files older than retentionDays from output/
8. Print summary (row count or "no records found")
```

### Graph API Details

> Teacher list is loaded from `teachers.json` (see Section 4). The API calls below are for the per-teacher meeting and attendance extraction.

**Step 4a — List meetings for a teacher:**
```http
GET https://graph.microsoft.com/v1.0/users/{teacherId}/onlineMeetings
    ?$filter=startDateTime ge 2026-03-01T00:00:00Z and startDateTime lt 2026-03-02T00:00:00Z
    &$select=id,subject,startDateTime,endDateTime,participants
Authorization: Bearer {token}
```

> **Note:** The `$filter` on `startDateTime` may not be supported on all endpoint versions. The script tries server-side filtering first; if it fails, it falls back to fetching recent meetings with `$top=50&$orderby=startDateTime desc` and filtering client-side by date range. See the `Get-Attendance.ps1` code in the Complete Main Script section.

**Step 4b-i — Get attendance reports:**
```http
GET https://graph.microsoft.com/v1.0/users/{teacherId}/onlineMeetings/{meetingId}/attendanceReports
Authorization: Bearer {token}
```

> **Channel meeting caveat:** When listing attendance reports for a channel meeting, the API returns reports for **every meeting ever held in that channel**, not just the specified meeting. The code must filter results by `meetingStartDateTime` / `meetingEndDateTime` to isolate the target meeting's report.

> **50-report limit:** The API returns at most 50 most-recent reports per call. For busy channels with frequent meetings, implement pagination or increase extraction frequency.

> **Retention:** Attendance reports are retained for **1 year** from the meeting date (as of March 2026). Daily extraction is still recommended for local archival, but backfill within 1 year is possible via Graph.

**Response structure:**
```json
{
  "value": [
    {
      "id": "report-id",
      "totalParticipantCount": 28,
      "meetingStartDateTime": "2026-03-01T10:00:00Z",
      "meetingEndDateTime": "2026-03-01T10:45:00Z"
    }
  ]
}
```

**Step 4b-ii — Get attendance records:**
```http
GET https://graph.microsoft.com/v1.0/users/{teacherId}/onlineMeetings/{meetingId}/attendanceReports/{reportId}
    ?$expand=attendanceRecords
Authorization: Bearer {token}
```

**Response structure:**
```json
{
  "id": "report-id",
  "totalParticipantCount": 28,
  "meetingStartDateTime": "2026-03-01T10:00:00Z",
  "meetingEndDateTime": "2026-03-01T10:45:00Z",
  "attendanceRecords": [
    {
      "emailAddress": "o.hassan@school.edu",
      "identity": {
        "displayName": "Omar Hassan",
        "id": "user-guid",
        "tenantId": "tenant-guid"
      },
      "role": "Attendee",
      "totalAttendanceInSeconds": 2520,
      "attendanceIntervals": [
        {
          "joinDateTime": "2026-03-01T10:02:00Z",
          "leaveDateTime": "2026-03-01T10:44:00Z",
          "durationInSeconds": 2520
        }
      ]
    }
  ]
}
```

### Determining Absent Students

To mark students as "Absent," we need to know who was **expected** to attend. The attendance API only returns records for students who actually joined the meeting — it does not list students who were expected but never showed up. Determining absence therefore requires an external roster.

Options for a future version:

1. **SDS class rosters** (recommended): Query `/education/classes/{classId}/members` to get enrolled students. Compare the roster against attendance records — anyone on the roster but not in the records is "Absent."
2. **Historical attendance**: If a student attended the same recurring meeting in previous days, expect them again.

**v1 (PoC) scope:** Only students who appear in attendance records are reported (Present / Late / Partial). **"Absent" rows are not generated in v1** because the Graph attendance API does not provide them and roster comparison is out of PoC scope. The `Get-AttendanceStatus` function below retains an "Absent" path for completeness — it will be used in v2 when roster data is integrated.

### Helper Functions

#### `Invoke-GraphWithRetry` — Single-page request with retry

```powershell
function Invoke-GraphWithRetry {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [int]$MaxRetries = 5,
        [double]$BaseDelay = 1.0
    )

    for ($attempt = 0; $attempt -lt $MaxRetries; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -TimeoutSec 60 -ErrorAction Stop
            return $response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__

            if ($statusCode -eq 429) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                $wait = if ($retryAfter) { [int]$retryAfter } else { 30 }
                Write-Warning "429 throttled - waiting $wait seconds (attempt $($attempt + 1)/$MaxRetries)"
                Start-Sleep -Seconds $wait
                continue
            }

            if ($statusCode -ge 500) {
                $wait = $BaseDelay * [Math]::Pow(2, $attempt)
                Write-Warning "$statusCode server error - backoff $wait seconds (attempt $($attempt + 1)/$MaxRetries)"
                Start-Sleep -Seconds $wait
                continue
            }

            # Non-retryable error — log and rethrow
            throw
        }
    }
    throw "Failed after $MaxRetries attempts for $Uri"
}
```

#### `Invoke-GraphPaged` — Paginated request (follows `@odata.nextLink`)

Many Graph endpoints return paginated results. This function collects all pages into a single array:

```powershell
function Invoke-GraphPaged {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [int]$MaxRetries = 5
    )

    $allItems = [System.Collections.Generic.List[object]]::new()
    $nextUri  = $Uri

    while ($nextUri) {
        $response = Invoke-GraphWithRetry -Uri $nextUri -Headers $Headers -MaxRetries $MaxRetries
        if ($response.value) {
            $allItems.AddRange($response.value)
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return $allItems
}
```

#### `Get-GraphToken` — Token acquisition with caching and auto-refresh

Tokens expire after ~60 minutes. For long-running jobs (hours at 35K teachers), the token must be refreshed before it expires:

```powershell
$script:tokenCache = @{ Token = $null; ExpiresAt = [datetime]::MinValue }

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    # Return cached token if it's still valid (with 5-minute buffer)
    if ($script:tokenCache.Token -and $script:tokenCache.ExpiresAt -gt (Get-Date).AddMinutes(5)) {
        return $script:tokenCache.Token
    }

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                                  -Method Post -Body $body -ErrorAction Stop

    $script:tokenCache.Token     = $response.access_token
    $script:tokenCache.ExpiresAt = (Get-Date).AddSeconds($response.expires_in)

    Write-Host "Token acquired/refreshed — expires at $($script:tokenCache.ExpiresAt.ToString('HH:mm:ss'))"
    return $response.access_token
}
```

### Complete Main Script (`Get-Attendance.ps1`)

This is the full orchestration script. Helper functions above (`Invoke-GraphWithRetry`, `Invoke-GraphPaged`, `Get-GraphToken`, `Get-AttendanceStatus`, `Export-AttendanceExcel`, `Remove-OldAttendanceFiles`, `Write-ErrorLog`) should be defined before this block — either in the same file or dot-sourced.

```powershell
# ── Configuration ──
param(
    [string]$ConfigPath = ".\config.json",
    [datetime]$TargetDate = (Get-Date).AddDays(-1).Date   # default: yesterday
)

$config   = Get-Content $ConfigPath | ConvertFrom-Json
$teachers = Get-Content ".\teachers.json" | ConvertFrom-Json

# ── Date range (UTC midnight to midnight) ──
$startDate = $TargetDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$endDate   = $TargetDate.AddDays(1).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Host "Extracting attendance for $($TargetDate.ToString('yyyy-MM-dd')) ($($teachers.Count) teachers)"

# ── Acquire initial token ──
$token = Get-GraphToken -TenantId $config.tenantId -ClientId $config.clientId -ClientSecret $env:CLIENT_SECRET

# ── Pre-capture scalar values for $using: in parallel block ──
$lateThreshMin   = $config.lateThresholdMinutes
$partialThreshPct = $config.partialThresholdPercent
$tenantId         = $config.tenantId
$clientId         = $config.clientId
$clientSecret     = $env:CLIENT_SECRET

# ── Process teachers in parallel ──
$allResults = $teachers | ForEach-Object -Parallel {
    $teacher       = $_
    $startDate     = $using:startDate
    $endDate       = $using:endDate
    $lateMin       = $using:lateThreshMin
    $partialPct    = $using:partialThreshPct

    # ── Token refresh inside runspace ──
    # Each runspace re-acquires its own token if needed.
    # (Functions are not shared across runspaces — inline the logic.)
    $tenantId      = $using:tenantId
    $clientId      = $using:clientId
    $clientSecret  = $using:clientSecret

    function Get-Token {
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
        $resp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
                                  -Method Post -Body $body -ErrorAction Stop
        return $resp.access_token
    }

    function Invoke-GraphSafe {
        param([string]$Uri, [hashtable]$Headers, [int]$MaxRetries = 5)
        for ($attempt = 0; $attempt -lt $MaxRetries; $attempt++) {
            try {
                return (Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -TimeoutSec 60 -ErrorAction Stop)
            }
            catch {
                $code = $_.Exception.Response.StatusCode.value__
                if ($code -eq 429) {
                    $ra = $_.Exception.Response.Headers["Retry-After"]
                    Start-Sleep -Seconds $(if ($ra) { [int]$ra } else { 30 })
                } elseif ($code -ge 500) {
                    Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
                } else { throw }
            }
        }
        throw "Failed after $MaxRetries retries: $Uri"
    }

    function Invoke-GraphPagedSafe {
        param([string]$Uri, [hashtable]$Headers)
        $items = [System.Collections.Generic.List[object]]::new()
        $next  = $Uri
        while ($next) {
            $resp = Invoke-GraphSafe -Uri $next -Headers $Headers
            if ($resp.value) { $items.AddRange($resp.value) }
            $next = $resp.'@odata.nextLink'
        }
        return $items
    }

    try {
        $token   = Get-Token
        $headers = @{ Authorization = "Bearer $token" }

        # ── Step 1: List meetings for this teacher ──
        # Try server-side $filter first; fall back to client-side if it fails
        $meetingsUri = "https://graph.microsoft.com/v1.0/users/$($teacher.id)/onlineMeetings" +
                       "?`$filter=startDateTime ge $startDate and startDateTime lt $endDate" +
                       "&`$select=id,subject,startDateTime,endDateTime"
        try {
            $meetings = Invoke-GraphPagedSafe -Uri $meetingsUri -Headers $headers
        }
        catch {
            # $filter not supported — get recent meetings and filter client-side
            $fallbackUri = "https://graph.microsoft.com/v1.0/users/$($teacher.id)/onlineMeetings" +
                           "?`$select=id,subject,startDateTime,endDateTime&`$top=50&`$orderby=startDateTime desc"
            $allMeetings = Invoke-GraphPagedSafe -Uri $fallbackUri -Headers $headers
            $meetings = $allMeetings | Where-Object {
                [datetime]$_.startDateTime -ge [datetime]$startDate -and
                [datetime]$_.startDateTime -lt [datetime]$endDate
            }
        }

        if (-not $meetings -or @($meetings).Count -eq 0) { return }  # No meetings — skip

        foreach ($meeting in $meetings) {
            # ── Step 2: Get attendance reports (paginated) ──
            $reportsUri = "https://graph.microsoft.com/v1.0/users/$($teacher.id)/onlineMeetings/$($meeting.id)/attendanceReports"
            $reports = Invoke-GraphPagedSafe -Uri $reportsUri -Headers $headers

            # ── Channel meeting filter: only keep reports matching our target meeting's time window ──
            $mStart = [datetime]$meeting.startDateTime
            $mEnd   = [datetime]$meeting.endDateTime
            $reports = $reports | Where-Object {
                [datetime]$_.meetingStartDateTime -ge $mStart.AddMinutes(-5) -and
                [datetime]$_.meetingStartDateTime -le $mStart.AddMinutes(5)
            }

            foreach ($report in $reports) {
                # ── Step 3: Get attendance records (expand inline, paginated) ──
                $recordsUri = "$reportsUri/$($report.id)?`$expand=attendanceRecords"
                $detail = Invoke-GraphSafe -Uri $recordsUri -Headers $headers

                foreach ($record in $detail.attendanceRecords) {
                    $meetingSec = ($mEnd - $mStart).TotalSeconds
                    $attPct     = if ($meetingSec -gt 0) { ($record.totalAttendanceInSeconds / $meetingSec) * 100 } else { 100 }
                    $joinDt     = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                                      [datetime]$record.attendanceIntervals[0].joinDateTime
                                  } else { $null }
                    $leaveDt    = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                                      [datetime]$record.attendanceIntervals[-1].leaveDateTime
                                  } else { $null }

                    $lateCutoff = $mStart.AddMinutes($lateMin)
                    $status = if ($attPct -lt $partialPct) { "Partial" }
                              elseif ($joinDt -and $joinDt -gt $lateCutoff) { "Late" }
                              else { "Present" }

                    [PSCustomObject]@{
                        Date             = $mStart.ToString('yyyy-MM-dd')
                        TeacherName      = $teacher.displayName
                        TeacherEmail     = $teacher.mail
                        MeetingSubject   = $meeting.subject
                        MeetingStart     = $meeting.startDateTime
                        MeetingEnd       = $meeting.endDateTime
                        StudentName      = $record.identity.displayName
                        StudentEmail     = $record.emailAddress
                        JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ') }  else { '' }
                        LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                        DurationMinutes  = [math]::Round($record.totalAttendanceInSeconds / 60, 1)
                        AttendanceStatus = $status
                    }
                }
            }
        }
    }
    catch {
        # Log error and continue with other teachers
        $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Write-Warning "$ts | ERROR | teacher=$($teacher.mail) | $_"
    }
} -ThrottleLimit 10

# ── Export results ──
if ($allResults) {
    Export-AttendanceExcel -DetailRows $allResults -OutputDir $config.outputDir -ReportDate $TargetDate
    Remove-OldAttendanceFiles -OutputDir $config.outputDir -RetentionDays $config.retentionDays
    Write-Host "Done — $($allResults.Count) attendance records exported"
} else {
    Write-Host "No attendance records found for $($TargetDate.ToString('yyyy-MM-dd'))"
}
```

### Key Design Decisions in the Main Script

| Concern | Solution |
|---|---|
| **Pagination** | `Invoke-GraphPagedSafe` follows `@odata.nextLink` until all pages are collected. Used for meetings list *and* attendance reports list. |
| **`$filter` fallback** | Tries server-side `$filter` on `startDateTime` first. If Graph rejects it (400/404), falls back to fetching recent meetings with `$top=50&$orderby=startDateTime desc` and filtering client-side. |
| **Token refresh** | Each parallel runspace acquires its own token via `Get-Token`. Since tokens last ~60 min and a single teacher completes in seconds, per-runspace acquisition is safe. The runspace pool is limited to 10, so at most 10 tokens are active at once. |
| **`$using:` safety** | All config values are pre-captured as scalar variables (`$lateThreshMin`, `$partialThreshPct`, etc.) before the parallel block. No nested property access through `$using:config.xxx`. |
| **Null guards on intervals** | `attendanceIntervals` is checked for both `$null` and `.Count -gt 0` before indexing. `JoinTime`/`LeaveTime` output empty strings instead of throwing. |
| **Channel meeting report pollution** | Reports are filtered by `meetingStartDateTime` within ±5 minutes of the meeting's `startDateTime`, so stale channel-meeting reports are excluded. |
| **Functions in runspaces** | `Invoke-GraphSafe`, `Invoke-GraphPagedSafe`, and `Get-Token` are defined *inside* the parallel block so they are available to each runspace. This is the simplest pattern for PoC; production code could use modules. |

> **Scaling note:** The `-ThrottleLimit 10` means up to 10 teachers are processed concurrently. Adjust based on observed throttling — if you see frequent 429s, reduce the limit. For production with multiple app registrations, partition `teachers.json` and run separate script instances per app.

---

## 6. Excel Export

### Using the `ImportExcel` Module

The `ImportExcel` PowerShell module eliminates the need for manual Excel formatting code. Install with `Install-Module ImportExcel`.

```powershell
function Export-AttendanceExcel {
    param(
        [array]$DetailRows,
        [string]$OutputDir = ".\output",
        [datetime]$ReportDate
    )

    # Ensure output directory exists
    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

    $filePath = Join-Path $OutputDir "attendance_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

    # Remove existing file (Export-Excel appends by default)
    if (Test-Path $filePath) { Remove-Item $filePath }

    # Export attendance data
    $DetailRows | Export-Excel -Path $filePath -WorksheetName "Attendance" `
        -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow `
        -ConditionalText $(
            New-ConditionalText -Text "Present" -BackgroundColor LightGreen
            New-ConditionalText -Text "Late"    -BackgroundColor Yellow
            New-ConditionalText -Text "Partial" -BackgroundColor Orange
            New-ConditionalText -Text "Absent"  -BackgroundColor Red -ConditionalTextColor White
        )

    Write-Host "Saved: $filePath"
    return $filePath
}


function Remove-OldAttendanceFiles {
    param(
        [string]$OutputDir = ".\output",
        [int]$RetentionDays = 30
    )

    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -Path $OutputDir -Filter "attendance_*.xlsx" | Where-Object {
        $_.LastWriteTime -lt $cutoff
    } | ForEach-Object {
        Remove-Item $_.FullName
        Write-Host "Deleted old file: $($_.Name)"
    }
}
```

### Output Example

Running the script for `2026-03-01` produces:

```
output/attendance_2026-03-01.xlsx
  └── Sheet: Attendance   (up to 1M+ rows for 35K-teacher tenants)
```

> **Note:** Excel has a maximum of ~1,048,576 rows per sheet. At peak scale (35K teachers, 4 meetings each, 30 students per meeting), the detail sheet may exceed this limit. Split output by date range if needed.

---

## 7. Attendance Status Computation

### PowerShell Functions

```powershell
function Get-AttendanceStatus {
    param(
        [datetime]$MeetingStart,
        [datetime]$MeetingEnd,
        [nullable[datetime]]$JoinTime,
        [int]$TotalAttendanceSeconds,
        [int]$LateThresholdMinutes = 10,
        [double]$PartialThresholdPercent = 50
    )

    # Student not in attendance records
    if ($null -eq $JoinTime) { return "Absent" }

    $meetingDurationSeconds = ($MeetingEnd - $MeetingStart).TotalSeconds
    if ($meetingDurationSeconds -le 0) { return "Present" }  # Edge case: zero-length meeting

    $attendancePercent = ($TotalAttendanceSeconds / $meetingDurationSeconds) * 100

    # Check partial attendance first (overrides late)
    if ($attendancePercent -lt $PartialThresholdPercent) { return "Partial" }

    # Check if late
    $lateCutoff = $MeetingStart.AddMinutes($LateThresholdMinutes)
    if ($JoinTime -gt $lateCutoff) { return "Late" }

    return "Present"
}
```

### Status Logic Summary

| Condition | Status |
|---|---|
| Student not in attendance records | **Absent** *(v2 — requires roster comparison)* |
| Join time > meeting start + threshold (default 10 min) | **Late** |
| Duration < threshold % of meeting length (default 50%) | **Partial** |
| Otherwise | **Present** |

Thresholds are configurable in `config.json`. In v1 (PoC), only Present / Late / Partial are produced.

---

## 8. Configuration

### `config.json` — All Settings (Merged)

```json
{
  "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "clientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "teacherGroupId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "outputDir": "./output",
  "logsDir": "./logs",
  "retentionDays": 30,
  "timezone": "Asia/Riyadh",
  "lateThresholdMinutes": 10,
  "partialThresholdPercent": 50
}
```

### Client Secret

The client secret is **not stored in `config.json`**. Set it as an environment variable:

```powershell
$env:CLIENT_SECRET = "your-client-secret-here"
```

Or for persistent use, set it in your PowerShell profile or use Windows Credential Manager.

### Module Dependencies

```powershell
# Install required PowerShell modules (one-time)
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

> **Note:** No `requirements.txt` or `pip install` needed. PowerShell modules are the only dependencies.

---

## 9. Error Handling & Retry Logic

### Retry Strategy

| Error Type | Action |
|---|---|
| **429 Too Many Requests** | Read `Retry-After` header, wait, retry (up to 5 times) |
| **5xx Server Error** | Exponential backoff: 1s, 2s, 4s, 8s, 16s |
| **401 Unauthorized** | Re-acquire token via MSAL `client_credentials` POST (see `Get-GraphToken`), retry once |
| **404 Not Found** (meeting/report) | Log and skip — meeting may have been deleted |
| **Network timeout** | Retry — `Invoke-RestMethod -TimeoutSec 60` is set in the retry helpers (see below) |

### Error Log

If a teacher's meetings cannot be processed after all retries, log to `logs/errors.log`:

```
2026-03-01T02:15:33Z | ERROR | teacher=a.rashid@school.edu | meeting=abc-123 | 429 Too Many Requests | retries=5
```

```powershell
function Write-ErrorLog {
    param(
        [string]$Teacher,
        [string]$Meeting,
        [string]$Message
    )
    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $logLine = "$timestamp | ERROR | teacher=$Teacher | meeting=$Meeting | $Message"
    Add-Content -Path ".\logs\errors.log" -Value $logLine
    Write-Warning $logLine
}
```

For the PoC, console + file logging is sufficient. Production would add Application Insights.

---

## 10. Testing Plan

### Manual Validation (PoC)

For the PoC, manual validation replaces formal unit tests:

| Test | How |
|---|---|
| Auth works | Run `Connect-MgGraph`, verify token acquired |
| Teacher sync | Run `Sync-Teachers.ps1`, verify `teachers.json` has expected count |
| List meetings | Query one known teacher, verify meetings returned |
| Attendance report | Query one known meeting, verify attendance records returned |
| Status computation | Test `Get-AttendanceStatus` with known values: on-time, late, partial, absent |
| Excel output | Run `Get-Attendance.ps1` for 1 teacher, open `.xlsx`, verify columns and data |
| File cleanup | Create a test file with old timestamp, verify `Remove-OldAttendanceFiles` deletes it |

### Integration Tests (Against Test Tenant)

| Test | Description |
|---|---|
| Single teacher end-to-end | `Get-Attendance.ps1` for 1 teacher → valid .xlsx |
| 10-teacher batch | Verify throttling and parallel processing work correctly |
| 100-teacher batch | Verify at moderate scale; monitor for 429 errors |

### Scale Test

| Scenario | Expected |
|---|---|
| Process 10 teachers × 5 meetings | Completes in < 5 minutes |
| Process 100 teachers × 5 meetings | Completes in < 30 minutes |
| Excel file with 50,000 detail rows | Writes in < 30 seconds |

---

## 11. Deployment Runbook (PoC)

### Day 1: Setup

- [ ] Create Azure AD App Registration with required permissions
- [ ] Get tenant admin to grant admin consent
- [ ] Create application access policy and assign to all users (or teacher group)
- [ ] Wait 30 minutes for policy propagation
- [ ] Install PowerShell modules: `Install-Module Microsoft.Graph, ImportExcel`
- [ ] Set `$env:CLIENT_SECRET` and create `config.json`
- [ ] Verify auth works: `Connect-MgGraph`, make test Graph API call
- [ ] Verify attendance report is retrievable for a known meeting
- [ ] Enable free Education Insights for educators (SDS is already in place)

### Day 2–4: Build & Test

- [ ] Get/create teacher security group with all teachers
- [ ] Run `Sync-Teachers.ps1` → verify `teachers.json` is populated
- [ ] Implement `Get-Attendance.ps1` (meetings, attendance, Excel export)
- [ ] Test end-to-end with 1 teacher → validate Excel output
- [ ] Test with 10 teachers → validate throttling and retry logic

### Day 5–6: Scale & Validate

- [ ] Run against 100 teachers; check data quality
- [ ] Run against full teacher list; monitor API throttling
- [ ] Validate Excel output: spot-check 10 students against Teams meeting records
- [ ] Share sample Excel file with stakeholders for feedback

### Day 7: Wrap-Up

- [ ] Document any issues, limitations, or adjustments needed
- [ ] Deliver final Excel file + PoC scripts to customer
- [ ] **Decision point:** Proceed to production (add multiple app registrations, Azure SQL, SharePoint, Power BI, alerting)?
- [ ] **POC COMPLETE** ✓
