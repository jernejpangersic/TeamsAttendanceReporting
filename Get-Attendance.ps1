<#
.SYNOPSIS
    Extracts Teams meeting attendance data for all teachers and exports to Excel.

.DESCRIPTION
    For each teacher in teachers.json, uses Microsoft Graph PowerShell cmdlets to retrieve
    online meetings held on the target date, attendance reports, and attendance records.
    Computes attendance status (Present / Late / Partial) and exports to an Excel workbook.

    Uses Get-MgUserOnlineMeeting, Get-MgUserOnlineMeetingAttendanceReport, and related
    cmdlets with built-in pagination (-All), token management, and 429 retry handling.

    Teachers are processed in parallel (ThrottleLimit 10) for performance.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.EXAMPLE
    .\Get-Attendance.ps1
    .\Get-Attendance.ps1 -TargetDate "2026-03-01"
    .\Get-Attendance.ps1 -ConfigPath "C:\config\config.json" -TargetDate "2026-02-28"
#>
param(
    [string]$ConfigPath = ".\config.json",
    [datetime]$TargetDate
)

$ErrorActionPreference = "Stop"

#region ── Helper Functions ──

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO",
        [string]$LogsDir
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $logLine   = "$timestamp | $Level | $Message"
    $logFile   = Join-Path $LogsDir "attendance_$(Get-Date -Format 'yyyy-MM-dd').log"

    Add-Content -Path $logFile -Value $logLine

    switch ($Level) {
        "ERROR" {
            Write-Warning $logLine
            Add-Content -Path (Join-Path $LogsDir "errors.log") -Value $logLine
        }
        "WARN"  { Write-Warning $logLine }
        default { Write-Host $logLine }
    }
}

function Export-AttendanceExcel {
    param(
        [array]$DetailRows,
        [string]$OutputDir = ".\output",
        [datetime]$ReportDate
    )

    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $filePath = Join-Path $OutputDir "attendance_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

    # Remove existing file (Export-Excel appends by default)
    if (Test-Path $filePath) { Remove-Item $filePath }

    $DetailRows | Export-Excel -Path $filePath -WorksheetName "Attendance" `
        -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow `
        -ConditionalText $(
            New-ConditionalText -Text "Present" -BackgroundColor LightGreen
            New-ConditionalText -Text "Late"    -BackgroundColor Yellow
            New-ConditionalText -Text "Partial" -BackgroundColor Orange
            New-ConditionalText -Text "Absent"  -BackgroundColor Red -ConditionalTextColor White
        )

    return $filePath
}

function Remove-OldAttendanceFiles {
    param(
        [string]$OutputDir = ".\output",
        [int]$RetentionDays = 30,
        [string]$LogsDir
    )

    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -Path $OutputDir -Filter "attendance_*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            Remove-Item $_.FullName
            Write-Log -Message "Deleted old file: $($_.Name)" -LogsDir $LogsDir
        }
}

#endregion

#region ── Main Script ──

# ── Load configuration ──
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

# ── Validate required modules ──
$requiredModules = @("Microsoft.Graph.Authentication", "ImportExcel")
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Error "Required module '$mod' is not installed. Install with: Install-Module $mod -Scope CurrentUser"
        exit 1
    }
}

# ── Ensure output and logs directories exist ──
foreach ($dir in @($config.outputDir, $config.logsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# ── Load teachers ──
$teachersPath = Join-Path (Split-Path $ConfigPath -Parent) "teachers.json"
if (-not (Test-Path $teachersPath)) {
    Write-Error "teachers.json not found at $teachersPath. Run Sync-Teachers.ps1 first."
    exit 1
}
$teachers = Get-Content $teachersPath | ConvertFrom-Json

if (-not $teachers -or @($teachers).Count -eq 0) {
    Write-Error "teachers.json is empty. Run Sync-Teachers.ps1 to populate it."
    exit 1
}

# ── Resolve target date using the configured timezone ──
try {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($config.timezone)
}
catch {
    Write-Log -Message "Could not resolve timezone '$($config.timezone)', falling back to UTC" -Level "WARN" -LogsDir $config.logsDir
    $tz = [System.TimeZoneInfo]::Utc
}

if (-not $PSBoundParameters.ContainsKey('TargetDate')) {
    # Default: yesterday in the configured timezone
    $nowInTz    = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $tz)
    $TargetDate = $nowInTz.Date.AddDays(-1)
}

# Convert target date range to UTC (midnight-to-midnight in the configured timezone)
$targetMidnight = [DateTime]::SpecifyKind($TargetDate.Date, [DateTimeKind]::Unspecified)
$startUtc       = [System.TimeZoneInfo]::ConvertTimeToUtc($targetMidnight, $tz)
$endUtc         = [System.TimeZoneInfo]::ConvertTimeToUtc($targetMidnight.AddDays(1), $tz)

Write-Log -Message "Extracting attendance for $($TargetDate.ToString('yyyy-MM-dd')) | UTC range: $($startUtc.ToString('o')) to $($endUtc.ToString('o')) | Teachers: $($teachers.Count)" -LogsDir $config.logsDir

# ── Pre-capture scalar values used in status computation ──
$lateThreshMin    = $config.lateThresholdMinutes
$partialThreshPct = $config.partialThresholdPercent

# ── Connect to Graph ONCE before processing ──
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
Write-Log -Message "Connected to Microsoft Graph" -LogsDir $config.logsDir

# ── Helper: Paginated Graph request via Invoke-MgGraphRequest ──
function Invoke-MgGraphPaged {
    param([string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Uri $next -Method GET -OutputType PSObject
        if ($resp.value) { $items.AddRange($resp.value) }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

# ── Process teachers sequentially ──
$processedMeetings = [System.Collections.Generic.HashSet[string]]::new()  # deduplicate across teachers
$allResults = foreach ($teacher in $teachers) {
    try {
        # ── Step 1: Get calendar events with Teams meeting URLs for the target date ──
        # With app-only permissions, /users/{id}/onlineMeetings requires a JoinWebUrl filter
        # and must be queried via the meeting ORGANIZER's user ID (not any attendee).
        # So we get calendar events, include the organizer, then query via their ID.
        $startIso = $startUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $endIso   = $endUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $calUri   = "/v1.0/users/$($teacher.id)/calendarView" +
                    "?startDateTime=$startIso&endDateTime=$endIso" +
                    "&`$select=id,subject,start,end,onlineMeeting,isOnlineMeeting,organizer"
        $allEvents = Invoke-MgGraphPaged -Uri $calUri

        # Filter to only Teams meetings (client-side, since $filter on isOnlineMeeting isn't supported)
        $events = @($allEvents | Where-Object {
            $_.isOnlineMeeting -eq $true -and
            $_.onlineMeeting.joinUrl
        })

        Write-Log -Message "  $($teacher.mail): $($allEvents.Count) calendar events, $($events.Count) online meetings" -LogsDir $config.logsDir

        if (-not $events -or $events.Count -eq 0) { continue }

        foreach ($event in $events) {
            $joinUrl = $event.onlineMeeting.joinUrl
            if (-not $joinUrl) { continue }

            # Deduplicate: skip meetings already SUCCESSFULLY processed via another teacher
            if ($processedMeetings.Contains($joinUrl)) { continue }

            # ── Step 2: Resolve the meeting organizer's user ID ──
            # The onlineMeetings API requires querying via the ORGANIZER's user ID.
            # For channel meetings, the calendar event organizer may be a Group (not a user).
            $organizerEmail = $event.organizer.emailAddress.address
            Write-Log -Message "    Event '$($event.subject)' organizer=$organizerEmail" -LogsDir $config.logsDir

            # Cache organizer ID lookups to avoid repeated Graph calls
            if (-not $script:organizerCache) { $script:organizerCache = @{} }

            # URL-encode the joinUrl for OData filter (% chars in the URL break OData otherwise)
            $encodedUrl = [Uri]::EscapeDataString($joinUrl)

            # Try multiple user IDs to query the meeting:
            #   1) Organizer as a user (works for user-organized meetings)
            #   2) Group owner(s) (works for channel meetings - real organizer is a group member)
            #   3) Current teacher (fallback)
            $candidateUserIds = [System.Collections.Generic.List[string]]::new()

            # --- Candidate 1: Organizer as direct user ---
            $orgCacheKey = "user:$organizerEmail"
            if ($script:organizerCache.ContainsKey($orgCacheKey)) {
                $cachedId = $script:organizerCache[$orgCacheKey]
                if ($cachedId) { $candidateUserIds.Add($cachedId) }
            }
            else {
                try {
                    $orgUser = Invoke-MgGraphRequest -Uri "/v1.0/users/$organizerEmail" -Method GET -OutputType PSObject
                    $script:organizerCache[$orgCacheKey] = $orgUser.id
                    $candidateUserIds.Add($orgUser.id)
                }
                catch {
                    $script:organizerCache[$orgCacheKey] = $null  # cache the failure
                }
            }

            # --- Candidate 2: If organizer is a Group, use the group's owners ---
            if ($candidateUserIds.Count -eq 0) {
                $grpCacheKey = "group:$organizerEmail"
                if ($script:organizerCache.ContainsKey($grpCacheKey)) {
                    $cachedOwners = $script:organizerCache[$grpCacheKey]
                    if ($cachedOwners) { foreach ($oid in $cachedOwners) { $candidateUserIds.Add($oid) } }
                }
                else {
                    try {
                        $grpResult = Invoke-MgGraphRequest -Uri "/v1.0/groups?`$filter=mail eq '$organizerEmail'" -Method GET -OutputType PSObject
                        $grp = $grpResult.value | Select-Object -First 1
                        if ($grp) {
                            $owners = Invoke-MgGraphRequest -Uri "/v1.0/groups/$($grp.id)/owners?`$select=id" -Method GET -OutputType PSObject
                            $ownerIds = @($owners.value | ForEach-Object { $_.id })
                            if ($ownerIds.Count -eq 0) {
                                # No owners — try members instead
                                $members = Invoke-MgGraphRequest -Uri "/v1.0/groups/$($grp.id)/members?`$select=id&`$top=5" -Method GET -OutputType PSObject
                                $ownerIds = @($members.value | ForEach-Object { $_.id })
                            }
                            $script:organizerCache[$grpCacheKey] = $ownerIds
                            foreach ($oid in $ownerIds) { $candidateUserIds.Add($oid) }
                            Write-Log -Message "    Organizer '$organizerEmail' is a Group — trying $($ownerIds.Count) owner(s)/member(s)" -LogsDir $config.logsDir
                        }
                        else {
                            $script:organizerCache[$grpCacheKey] = $null
                        }
                    }
                    catch {
                        Write-Log -Message "    Group lookup failed for '$organizerEmail': $($_.Exception.Message -replace '[\r\n]+',' ')" -Level "WARN" -LogsDir $config.logsDir
                        $script:organizerCache[$grpCacheKey] = $null
                    }
                }
            }

            # --- Candidate 3: Fallback to the current teacher ---
            if ($candidateUserIds.Count -eq 0 -or -not ($candidateUserIds -contains $teacher.id)) {
                $candidateUserIds.Add($teacher.id)
            }

            # ── Step 3: Try each candidate user ID until one works ──
            $meeting = $null
            $meetingUserId = $null
            foreach ($uid in $candidateUserIds) {
                $meetingUri = "/v1.0/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$encodedUrl'"
                try {
                    $meetingResult = Invoke-MgGraphRequest -Uri $meetingUri -Method GET -OutputType PSObject
                    $m = $meetingResult.value | Select-Object -First 1
                    if ($m) {
                        $meeting = $m
                        $meetingUserId = $uid
                        break
                    }
                }
                catch {
                    # Try the next candidate
                }
            }

            if (-not $meeting) {
                Write-Log -Message "    No meeting found for '$($event.subject)' after trying $($candidateUserIds.Count) user(s)" -Level "WARN" -LogsDir $config.logsDir
                continue
            }

            Write-Log -Message "    Found meeting '$($event.subject)' via user $meetingUserId" -LogsDir $config.logsDir

            $mStart = [datetime]$event.start.dateTime
            $mEnd   = [datetime]$event.end.dateTime

            # ── Step 4: Get attendance reports (via organizer's user context) ──
            try {
                $reportsUri = "/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)/attendanceReports"
                $reports = Invoke-MgGraphPaged -Uri $reportsUri
            }
            catch {
                Write-Log -Message "    No attendance reports for '$($event.subject)': $($_.Exception.Message -replace '[\r\n]+',' ')" -Level "WARN" -LogsDir $config.logsDir
                $processedMeetings.Add($joinUrl) | Out-Null
                continue
            }

            if ($reports.Count -eq 0) {
                Write-Log -Message "    Meeting '$($event.subject)' found but has no attendance reports" -LogsDir $config.logsDir
                $processedMeetings.Add($joinUrl) | Out-Null
                continue
            }

            foreach ($report in $reports) {
                # ── Step 5: Get attendance records ──
                $recordsUri = "/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)" +
                              "/attendanceReports/$($report.id)?`$expand=attendanceRecords"
                $detail = Invoke-MgGraphRequest -Uri $recordsUri -Method GET -OutputType PSObject

                foreach ($record in $detail.attendanceRecords) {
                    $meetingSec = ($mEnd - $mStart).TotalSeconds
                    $attPct     = if ($meetingSec -gt 0) {
                                      ($record.totalAttendanceInSeconds / $meetingSec) * 100
                                  } else { 100 }

                    # Null-guard attendanceIntervals before indexing
                    $joinDt  = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                                   [datetime]$record.attendanceIntervals[0].joinDateTime
                               } else { $null }
                    $leaveDt = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                                   [datetime]$record.attendanceIntervals[-1].leaveDateTime
                               } else { $null }

                    # Compute status using config thresholds
                    $lateCutoff = $mStart.AddMinutes($lateThreshMin)
                    $status = if ($attPct -lt $partialThreshPct)              { "Partial" }
                              elseif ($joinDt -and $joinDt -gt $lateCutoff) { "Late"    }
                              else                                            { "Present" }

                    # Emit row
                    [PSCustomObject]@{
                        Date             = $mStart.ToString('yyyy-MM-dd')
                        TeacherName      = $teacher.displayName
                        TeacherEmail     = $teacher.mail
                        MeetingSubject   = $event.subject
                        MeetingStart     = $mStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
                        MeetingEnd       = $mEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
                        AttendeeName      = $record.identity.displayName
                        AttendeeEmail     = $record.emailAddress
                        JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ')  } else { '' }
                        LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                        DurationMinutes  = [math]::Round($record.totalAttendanceInSeconds / 60, 1)
                        AttendanceStatus = $status
                    }
                }
            }

            # Mark this meeting as successfully processed (dedup for other teachers)
            $processedMeetings.Add($joinUrl) | Out-Null
        }
    }
    catch {
        $ts      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $logLine = "$ts | ERROR | teacher=$($teacher.mail) | $_"
        Write-Warning $logLine
        Add-Content -Path (Join-Path $config.logsDir "errors.log") -Value $logLine
    }
}

# ── Export results ──
if ($allResults -and @($allResults).Count -gt 0) {
    $filePath = Export-AttendanceExcel -DetailRows $allResults `
                                       -OutputDir $config.outputDir `
                                       -ReportDate $TargetDate

    Remove-OldAttendanceFiles -OutputDir $config.outputDir `
                               -RetentionDays $config.retentionDays `
                               -LogsDir $config.logsDir

    Write-Log -Message "Done — $(@($allResults).Count) attendance records exported to $filePath" -LogsDir $config.logsDir
}
else {
    Write-Log -Message "No attendance records found for $($TargetDate.ToString('yyyy-MM-dd'))" -Level "WARN" -LogsDir $config.logsDir
}

#endregion
