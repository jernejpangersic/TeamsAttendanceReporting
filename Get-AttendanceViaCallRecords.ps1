<#
.SYNOPSIS
    Extracts Teams meeting attendance data using Call Records as the discovery layer.

.DESCRIPTION
    Instead of polling every teacher's calendar (35K calendarView calls), this script:
      1. Lists all Call Records for the target date range (1 paginated tenant-wide query)
      2. Filters to group calls (meetings) organized by teachers in teachers.json
      3. Resolves the meeting via the organizer's ID + joinWebUrl
      4. Pulls attendance reports and records for each meeting
      5. Exports to Excel

    This eliminates the per-teacher calendar round-trip for inactive teachers.
    On a light day (25K inactive teachers), this saves ~25K API calls.

    Requires the additional permission: CallRecords.Read.All (application).

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.EXAMPLE
    .\Get-AttendanceViaCallRecords.ps1
    .\Get-AttendanceViaCallRecords.ps1 -TargetDate "2026-03-01"
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
    $logFile   = Join-Path $LogsDir "callrecords_$(Get-Date -Format 'yyyy-MM-dd').log"

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

function Get-AttendanceStatus {
    param(
        [datetime]$MeetingStart,
        [datetime]$MeetingEnd,
        [nullable[datetime]]$JoinTime,
        [int]$TotalAttendanceSeconds,
        [int]$LateThresholdMinutes = 10,
        [double]$PartialThresholdPercent = 50
    )

    if ($null -eq $JoinTime) { return "Absent" }

    $meetingDurationSeconds = ($MeetingEnd - $MeetingStart).TotalSeconds
    if ($meetingDurationSeconds -le 0) { return "Present" }

    $attendancePercent = ($TotalAttendanceSeconds / $meetingDurationSeconds) * 100

    if ($attendancePercent -lt $PartialThresholdPercent) { return "Partial" }

    $lateCutoff = $MeetingStart.AddMinutes($LateThresholdMinutes)
    if ($JoinTime -gt $lateCutoff) { return "Late" }

    return "Present"
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

    $filePath = Join-Path $OutputDir "callrecords_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

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
    Get-ChildItem -Path $OutputDir -Filter "callrecords_*.xlsx" -ErrorAction SilentlyContinue |
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

# ── Load teachers and build lookup set ──
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

# Build HashSet of teacher IDs for fast lookup + dictionary for metadata
$teacherIdSet  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$teacherLookup = @{}
foreach ($t in $teachers) {
    $teacherIdSet.Add($t.id) | Out-Null
    $teacherLookup[$t.id] = $t
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
    $nowInTz    = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $tz)
    $TargetDate = $nowInTz.Date.AddDays(-1)
}

$targetMidnight = [DateTime]::SpecifyKind($TargetDate.Date, [DateTimeKind]::Unspecified)
$startUtc       = [System.TimeZoneInfo]::ConvertTimeToUtc($targetMidnight, $tz)
$endUtc         = [System.TimeZoneInfo]::ConvertTimeToUtc($targetMidnight.AddDays(1), $tz)

# Call Records API rejects filter dates in the future — cap endUtc at now
$nowUtc = [DateTime]::UtcNow
if ($endUtc -gt $nowUtc) { $endUtc = $nowUtc }

$startIso = $startUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
$endIso   = $endUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

Write-Log -Message "=== Call Records approach ===" -LogsDir $config.logsDir
Write-Log -Message "Target date: $($TargetDate.ToString('yyyy-MM-dd')) | UTC range: $startIso to $endIso | Teachers: $($teachers.Count)" -LogsDir $config.logsDir

# ── Pre-capture config scalars ──
$lateThreshMin    = $config.lateThresholdMinutes
$partialThreshPct = $config.partialThresholdPercent

# ── Connect to Graph ──
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
Write-Log -Message "Connected to Microsoft Graph" -LogsDir $config.logsDir

# ── Helper: Paginated Graph request ──
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

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 1: Discover meetings via Call Records
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 1: Listing call records for $($TargetDate.ToString('yyyy-MM-dd'))..." -LogsDir $config.logsDir

# Note: The list endpoint does NOT return participants — only the individual GET does.
# We request organizer here for the initial teacher-organized bucket.
$callRecordsUri = "/v1.0/communications/callRecords" +
                  "?`$filter=startDateTime ge $startIso and startDateTime lt $endIso" +
                  "&`$select=id,type,startDateTime,endDateTime,joinWebUrl,organizer"

$allCallRecords = Invoke-MgGraphPaged -Uri $callRecordsUri

Write-Log -Message "  Total call records returned: $($allCallRecords.Count)" -LogsDir $config.logsDir

# Filter to group calls (meetings) that have a joinWebUrl
$meetingRecords = @($allCallRecords | Where-Object {
    $_.type -eq 'groupCall' -and $_.joinWebUrl
})

Write-Log -Message "  Group calls with joinWebUrl: $($meetingRecords.Count)" -LogsDir $config.logsDir

# Bucket 1: Meetings organized by our teachers (direct organizer match)
$teacherMeetings = @($meetingRecords | Where-Object {
    $orgId = $_.organizer.user.id
    $orgId -and $teacherIdSet.Contains($orgId)
})

Write-Log -Message "  Organized by teachers in our list: $($teacherMeetings.Count)" -LogsDir $config.logsDir

# Bucket 2: Meetings where organizer is NOT a teacher (channel/group/external meetings).
# We need to check participants to see if teachers were involved, but the list endpoint
# doesn't return participants. For these, we fetch the full record individually.
$nonTeacherRecords = @($meetingRecords | Where-Object {
    $orgId = $_.organizer.user.id
    -not ($orgId -and $teacherIdSet.Contains($orgId))
})

$channelMeetings = [System.Collections.Generic.List[object]]::new()
$skippedCount    = 0

if ($nonTeacherRecords.Count -gt 0) {
    Write-Log -Message "  Fetching participants for $($nonTeacherRecords.Count) non-teacher-organized record(s)..." -LogsDir $config.logsDir

    foreach ($cr in $nonTeacherRecords) {
        try {
            # Fetch full record to get participants
            $full = Invoke-MgGraphRequest -Uri "/v1.0/communications/callRecords/$($cr.id)?`$select=id,type,startDateTime,endDateTime,joinWebUrl,organizer,participants" -Method GET -OutputType PSObject

            # Check if any participant is a teacher
            $teacherParticipants = @($full.participants | Where-Object {
                $_.user.id -and $teacherIdSet.Contains($_.user.id)
            })

            if ($teacherParticipants.Count -gt 0) {
                # Attach teacher participants to the record for use in Phase 2
                $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue $teacherParticipants -Force
                $channelMeetings.Add($cr)
            }
            else {
                # Fallback: Extract the Oid (meeting creator) from the joinWebUrl
                # Format: ...context=%7b%22Tid%22%3a%22{tenantId}%22%2c%22Oid%22%3a%22{userId}%22%7d
                $oidMatch = [regex]::Match($cr.joinWebUrl, 'Oid%22%3a%22([0-9a-f-]+)%22')
                if ($oidMatch.Success) {
                    $creatorId = $oidMatch.Groups[1].Value
                    if ($teacherIdSet.Contains($creatorId)) {
                        $fakeParticipant = [PSCustomObject]@{ user = [PSCustomObject]@{ id = $creatorId } }
                        $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue @($fakeParticipant) -Force
                        $channelMeetings.Add($cr)
                        Write-Log -Message "    Record $($cr.id): creator from joinUrl Oid = $($teacherLookup[$creatorId].mail)" -LogsDir $config.logsDir
                    }
                    else {
                        $skippedCount++
                    }
                }
                else {
                    $skippedCount++
                }
            }
        }
        catch {
            Write-Log -Message "    Failed to fetch record $($cr.id): $_" -Level "WARN" -LogsDir $config.logsDir
            $skippedCount++
        }
    }
}

Write-Log -Message "  Channel/external meetings with teacher involvement: $($channelMeetings.Count)" -LogsDir $config.logsDir
if ($skippedCount -gt 0) {
    Write-Log -Message "  Skipped $skippedCount meetings with no teacher involvement" -LogsDir $config.logsDir
}

# Combine both sets for processing
$allMeetingsToProcess = @($teacherMeetings) + @($channelMeetings)

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2: Get attendance data for each discovered meeting
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 2: Retrieving attendance for $($allMeetingsToProcess.Count) meetings..." -LogsDir $config.logsDir

$processedUrls  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$allResults      = [System.Collections.Generic.List[object]]::new()
$meetingsSuccess = 0
$meetingsFailed  = 0

foreach ($cr in $allMeetingsToProcess) {
    $joinUrl = $cr.joinWebUrl

    # Deduplicate — same joinWebUrl can appear in multiple call records (recurring meetings have
    # separate call records per occurrence, but same joinWebUrl; attendance reports may overlap)
    if ($processedUrls.Contains($joinUrl)) { continue }
    $processedUrls.Add($joinUrl) | Out-Null

    # Build candidate user IDs to try for meeting resolution.
    # The onlineMeetings API with JoinWebUrl filter must be queried via the actual meeting
    # organizer's user ID. For teacher-organized meetings that's straightforward.
    # For channel meetings, the organizer may be a group — we try teacher participants
    # and the organizer (if it's a user) in sequence until one resolves.
    $candidateUserIds = [System.Collections.Generic.List[string]]::new()
    $teacherForRow    = $null  # the teacher to attribute this meeting to

    $orgId = $cr.organizer.user.id
    if ($orgId -and $teacherIdSet.Contains($orgId)) {
        # Organizer IS a teacher — use them directly
        $candidateUserIds.Add($orgId)
        $teacherForRow = $teacherLookup[$orgId]
    }
    else {
        # Channel / external meeting — try organizer first (if it's a user), then teacher participants
        if ($orgId) {
            $candidateUserIds.Add($orgId)
        }

        # Use _teacherParticipants attached in Phase 1 (from individual GET or Oid extraction)
        foreach ($p in $cr._teacherParticipants) {
            $pid = $p.user.id
            if ($pid -and $teacherIdSet.Contains($pid) -and -not $candidateUserIds.Contains($pid)) {
                $candidateUserIds.Add($pid)
                # Use the first teacher participant for attribution
                if (-not $teacherForRow) {
                    $teacherForRow = $teacherLookup[$pid]
                }
            }
        }
    }

    if (-not $teacherForRow) {
        Write-Log -Message "  No teacher found for meeting joinUrl — skipping" -Level "WARN" -LogsDir $config.logsDir
        $meetingsFailed++
        continue
    }

    try {
        # ── Step A: Resolve meeting via candidate user IDs + joinWebUrl ──
        $encodedUrl = [Uri]::EscapeDataString($joinUrl)
        $meeting       = $null
        $meetingUserId = $null

        foreach ($uid in $candidateUserIds) {
            try {
                $meetingUri    = "/v1.0/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$encodedUrl'"
                $meetingResult = Invoke-MgGraphRequest -Uri $meetingUri -Method GET -OutputType PSObject
                $m = $meetingResult.value | Select-Object -First 1
                if ($m) {
                    $meeting       = $m
                    $meetingUserId = $uid
                    break
                }
            }
            catch {
                # Try the next candidate
            }
        }

        if (-not $meeting) {
            Write-Log -Message "  No meeting object found for '$($cr.joinWebUrl)' after trying $($candidateUserIds.Count) user(s)" -Level "WARN" -LogsDir $config.logsDir
            $meetingsFailed++
            continue
        }

        Write-Log -Message "  Found meeting '$($meeting.subject)' via user $meetingUserId (teacher=$($teacherForRow.mail))" -LogsDir $config.logsDir

        # Use times from the call record (actual start/end) rather than scheduled times
        $mStart = [datetime]$cr.startDateTime
        $mEnd   = [datetime]$cr.endDateTime

        # ── Step B: Get attendance reports (paginated) ──
        $reportsUri = "/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)/attendanceReports"
        $reports = Invoke-MgGraphPaged -Uri $reportsUri

        if ($reports.Count -eq 0) {
            Write-Log -Message "  No attendance reports for '$($meeting.subject)' (teacher=$($teacherForRow.mail))" -Level "WARN" -LogsDir $config.logsDir
            $meetingsFailed++
            continue
        }

        foreach ($report in $reports) {
            # ── Step C: Get attendance records ──
            $recordsUri = "/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)" +
                          "/attendanceReports/$($report.id)?`$expand=attendanceRecords"
            $detail = Invoke-MgGraphRequest -Uri $recordsUri -Method GET -OutputType PSObject

            foreach ($record in $detail.attendanceRecords) {
                $meetingSec = ($mEnd - $mStart).TotalSeconds
                $attPct     = if ($meetingSec -gt 0) {
                                  ($record.totalAttendanceInSeconds / $meetingSec) * 100
                              } else { 100 }

                $joinDt  = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                               [datetime]$record.attendanceIntervals[0].joinDateTime
                           } else { $null }
                $leaveDt = if ($record.attendanceIntervals -and $record.attendanceIntervals.Count -gt 0) {
                               [datetime]$record.attendanceIntervals[-1].leaveDateTime
                           } else { $null }

                $lateCutoff = $mStart.AddMinutes($lateThreshMin)
                $status = if ($attPct -lt $partialThreshPct)              { "Partial" }
                          elseif ($joinDt -and $joinDt -gt $lateCutoff)   { "Late"    }
                          else                                              { "Present" }

                $allResults.Add([PSCustomObject]@{
                    Date             = $mStart.ToString('yyyy-MM-dd')
                    TeacherName      = $teacherForRow.displayName
                    TeacherEmail     = $teacherForRow.mail
                    MeetingSubject   = $meeting.subject
                    MeetingStart     = $mStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
                    MeetingEnd       = $mEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
                    StudentName      = $record.identity.displayName
                    StudentEmail     = $record.emailAddress
                    JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ')  } else { '' }
                    LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                    DurationMinutes  = [math]::Round($record.totalAttendanceInSeconds / 60, 1)
                    AttendanceStatus = $status
                })
            }
        }

        $meetingsSuccess++
        Write-Log -Message "  [$meetingsSuccess] $($meeting.subject) (teacher=$($teacherForRow.mail)) — $($reports.Count) report(s)" -LogsDir $config.logsDir
    }
    catch {
        $meetingsFailed++
        Write-Log -Message "Failed processing meeting for $($teacherForRow.mail): $_" -Level "ERROR" -LogsDir $config.logsDir
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 3: Export results
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 3: Export" -LogsDir $config.logsDir
Write-Log -Message "  Meetings processed: $meetingsSuccess success, $meetingsFailed failed/skipped" -LogsDir $config.logsDir
Write-Log -Message "  Unique joinUrls seen: $($processedUrls.Count)" -LogsDir $config.logsDir

if ($allResults.Count -gt 0) {
    $filePath = Export-AttendanceExcel -DetailRows $allResults `
                                       -OutputDir $config.outputDir `
                                       -ReportDate $TargetDate

    Remove-OldAttendanceFiles -OutputDir $config.outputDir `
                               -RetentionDays $config.retentionDays `
                               -LogsDir $config.logsDir

    Write-Log -Message "Done — $($allResults.Count) attendance records exported to $filePath" -LogsDir $config.logsDir
}
else {
    Write-Log -Message "No attendance records found for $($TargetDate.ToString('yyyy-MM-dd'))" -Level "WARN" -LogsDir $config.logsDir
}

Disconnect-MgGraph -ErrorAction SilentlyContinue

#endregion
