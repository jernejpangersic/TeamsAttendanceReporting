<#
.SYNOPSIS
    Extracts Teams meeting attendance data using Call Records — with parallel Phase 2.

.DESCRIPTION
    Same logic as Get-AttendanceViaCallRecords.ps1, but Phase 2 (meeting resolution
    and attendance fetching) runs in parallel using ForEach-Object -Parallel (PS 7+).

    Phase 2 is the dominant bottleneck: each meeting requires 2-4 sequential HTTP
    round-trips. By processing meetings concurrently (default: 8 threads), wall-clock
    time drops roughly proportional to the throttle limit.

    The parallel workers use Invoke-RestMethod with a pre-acquired OAuth token instead
    of Invoke-MgGraphRequest, because the MgGraph session isn't available inside
    parallel runspaces. Retry-after logic handles Graph API 429 throttling.

    Requires PowerShell 7+ for ForEach-Object -Parallel.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.PARAMETER ThrottleLimit
    Maximum concurrent threads for Phase 2 Graph API calls. Default: 8.
    Higher values are faster but increase risk of Graph 429 throttling.

.EXAMPLE
    .\Get-AttendanceViaCallRecords-Parallel.ps1
    .\Get-AttendanceViaCallRecords-Parallel.ps1 -TargetDate "2026-03-01" -ThrottleLimit 12
#>
param(
    [string]$ConfigPath = ".\config.json",
    [datetime]$TargetDate,
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 8
)

$ErrorActionPreference = "Stop"

# ── Require PowerShell 7+ for ForEach-Object -Parallel ──
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7+ for parallel execution. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}

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
    $logFile   = Join-Path $LogsDir "callrecords_parallel_$(Get-Date -Format 'yyyy-MM-dd').log"

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

    $filePath = Join-Path $OutputDir "callrecords_parallel_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

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
    Get-ChildItem -Path $OutputDir -Filter "callrecords_parallel_*.xlsx" -ErrorAction SilentlyContinue |
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

Write-Log -Message "=== Call Records approach (Parallel Phase 2, ThrottleLimit=$ThrottleLimit) ===" -LogsDir $config.logsDir
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
# PHASE 1: Discover meetings via Call Records  (unchanged — serial, 1 query)
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 1: Listing call records for $($TargetDate.ToString('yyyy-MM-dd'))..." -LogsDir $config.logsDir

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

# Bucket 2: Meetings where organizer is NOT a teacher
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
            $full = Invoke-MgGraphRequest -Uri "/v1.0/communications/callRecords/$($cr.id)?`$select=id,type,startDateTime,endDateTime,joinWebUrl,organizer,participants" -Method GET -OutputType PSObject

            $teacherParticipants = @($full.participants | Where-Object {
                $_.user.id -and $teacherIdSet.Contains($_.user.id)
            })

            if ($teacherParticipants.Count -gt 0) {
                $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue $teacherParticipants -Force
                $channelMeetings.Add($cr)
            }
            else {
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

$allMeetingsToProcess = @($teacherMeetings) + @($channelMeetings)

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2: Get attendance data — PARALLEL
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 2: Retrieving attendance for $($allMeetingsToProcess.Count) meetings (parallel, ThrottleLimit=$ThrottleLimit)..." -LogsDir $config.logsDir

# ── Pre-deduplicate by joinWebUrl (must happen before parallel dispatch) ──
$uniqueMeetings = [System.Collections.Generic.List[object]]::new()
$seenUrls       = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($cr in $allMeetingsToProcess) {
    if ($cr.joinWebUrl -and $seenUrls.Add($cr.joinWebUrl)) {
        $uniqueMeetings.Add($cr)
    }
}

Write-Log -Message "  Unique meetings after dedup: $($uniqueMeetings.Count)" -LogsDir $config.logsDir

# ── Pre-compute work items (no API calls, just local data prep) ──
# This avoids passing complex lookup structures into parallel runspaces.
$workItems = foreach ($cr in $uniqueMeetings) {
    $candidateUserIds = [System.Collections.Generic.List[string]]::new()
    $teacherForRow    = $null

    $orgId = $cr.organizer.user.id
    if ($orgId -and $teacherIdSet.Contains($orgId)) {
        $candidateUserIds.Add($orgId)
        $teacherForRow = $teacherLookup[$orgId]
    }
    else {
        if ($orgId) { $candidateUserIds.Add($orgId) }

        foreach ($p in $cr._teacherParticipants) {
            $pid2 = $p.user.id
            if ($pid2 -and $teacherIdSet.Contains($pid2) -and -not $candidateUserIds.Contains($pid2)) {
                $candidateUserIds.Add($pid2)
                if (-not $teacherForRow) {
                    $teacherForRow = $teacherLookup[$pid2]
                }
            }
        }
    }

    if (-not $teacherForRow) { continue }

    # Emit a serializable work item
    [PSCustomObject]@{
        JoinWebUrl       = $cr.joinWebUrl
        StartDateTime    = $cr.startDateTime
        EndDateTime      = $cr.endDateTime
        CandidateUserIds = @($candidateUserIds)
        TeacherName      = $teacherForRow.displayName
        TeacherEmail     = $teacherForRow.mail
        TeacherDept      = $teacherForRow.department
    }
}

$workItems = @($workItems)

if ($workItems.Count -eq 0) {
    Write-Log -Message "  No meetings to process after work-item preparation" -Level "WARN" -LogsDir $config.logsDir
}
else {
    Write-Log -Message "  Dispatching $($workItems.Count) work items to $ThrottleLimit parallel workers..." -LogsDir $config.logsDir
}

# ── Acquire a raw OAuth token for parallel workers ──
# Parallel runspaces don't share the MgGraph session, so we use Invoke-RestMethod
# with a bearer token instead.
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $config.clientId
    client_secret = $config.clientSecret
    scope         = "https://graph.microsoft.com/.default"
}
$tokenResponse = Invoke-RestMethod `
    -Uri "https://login.microsoftonline.com/$($config.tenantId)/oauth2/v2.0/token" `
    -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
$accessToken = $tokenResponse.access_token

if (-not $accessToken) {
    Write-Error "Failed to acquire OAuth access token for parallel workers."
    exit 1
}

# ── Thread-safe progress state shared across parallel workers ──
$totalWorkItems = $workItems.Count
$sharedProgress = [hashtable]::Synchronized(@{ Done = 0; Errors = 0; Throttled = 0 })
# Log a text line every ~10% of total, but at least every 50 and at most every 500
$logInterval    = [math]::Max(50, [math]::Min(500, [math]::Ceiling($totalWorkItems / 10)))

# ── Parallel execution ──
$phase2Timer = [System.Diagnostics.Stopwatch]::StartNew()

$parallelOutput = $workItems | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $wi               = $_
    $token            = $using:accessToken
    $lateThreshMin    = $using:lateThreshMin
    $partialThreshPct = $using:partialThreshPct
    $graphBase        = "https://graph.microsoft.com"

    # Import shared progress state
    $progress         = $using:sharedProgress
    $total            = $using:totalWorkItems
    $logEvery         = $using:logInterval

    $headers = @{ Authorization = "Bearer $token" }

    # ── Helper: Invoke-RestMethod with retry on 429 / 503 / 504 ──
    function Invoke-GraphRest {
        param(
            [string]$Uri,
            [hashtable]$Headers,
            [int]$MaxRetries = 5
        )
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method GET
            }
            catch {
                $response = $_.Exception.Response
                $status   = if ($response) { [int]$response.StatusCode } else { 0 }

                # Retry on throttling (429) or transient server errors (503, 504)
                $retryable = $status -in @(429, 503, 504)

                if ($retryable -and $attempt -lt $MaxRetries) {
                    # Read Retry-After header; fall back to exponential backoff
                    $waitSec = [math]::Pow(2, $attempt)   # default: 2, 4, 8, 16, 32s
                    if ($response -and $response.Headers) {
                        try {
                            $raValues = $null
                            if ($response.Headers.TryGetValues('Retry-After', [ref]$raValues)) {
                                $parsed = [int]($raValues | Select-Object -First 1)
                                if ($parsed -gt 0) { $waitSec = $parsed }
                            }
                        } catch { }
                    }

                    if ($status -eq 429) { $progress.Throttled++ }
                    $logs.Add("WARN | Throttled (HTTP $status) on attempt $attempt/$MaxRetries — waiting ${waitSec}s before retry: $Uri")
                    Start-Sleep -Seconds $waitSec
                }
                else {
                    if ($retryable) {
                        $logs.Add("ERROR | Exhausted $MaxRetries retries (HTTP $status) for: $Uri")
                    }
                    throw
                }
            }
        }
    }

    # ── Helper: Paginated GET ──
    function Invoke-GraphPaged {
        param([string]$Uri, [hashtable]$Headers)
        $items = [System.Collections.Generic.List[object]]::new()
        $next  = $Uri
        while ($next) {
            $resp = Invoke-GraphRest -Uri $next -Headers $Headers
            if ($resp.value) { $items.AddRange(@($resp.value)) }
            $next = $resp.'@odata.nextLink'
        }
        return $items
    }

    # ── Per-meeting processing ──
    $results = [System.Collections.Generic.List[object]]::new()
    $logs    = [System.Collections.Generic.List[string]]::new()
    $success = $false

    try {
        # Step A: Resolve meeting via candidate user IDs + joinWebUrl
        $encodedUrl = [Uri]::EscapeDataString($wi.JoinWebUrl)
        $meeting       = $null
        $meetingUserId = $null

        foreach ($uid in $wi.CandidateUserIds) {
            try {
                $meetingUri = "$graphBase/v1.0/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$encodedUrl'"
                $meetingResult = Invoke-GraphRest -Uri $meetingUri -Headers $headers
                $m = $meetingResult.value | Select-Object -First 1
                if ($m) {
                    $meeting       = $m
                    $meetingUserId = $uid
                    break
                }
            }
            catch {
                # Try next candidate
            }
        }

        if (-not $meeting) {
            $logs.Add("WARN | No meeting object found for '$($wi.JoinWebUrl)' after trying $($wi.CandidateUserIds.Count) user(s)")
            # Return partial result
            [PSCustomObject]@{
                Results = @()
                Logs    = $logs.ToArray()
                Success = $false
            }
            return
        }

        $logs.Add("INFO | Found meeting '$($meeting.subject)' via user $meetingUserId (teacher=$($wi.TeacherEmail))")

        # Use times from the call record (actual start/end)
        $mStart = [datetime]$wi.StartDateTime
        $mEnd   = [datetime]$wi.EndDateTime

        # Step B: Get attendance reports
        $reportsUri = "$graphBase/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)/attendanceReports"
        $reports = Invoke-GraphPaged -Uri $reportsUri -Headers $headers

        if ($reports.Count -eq 0) {
            $logs.Add("WARN | No attendance reports for '$($meeting.subject)' (teacher=$($wi.TeacherEmail))")
            [PSCustomObject]@{
                Results = @()
                Logs    = $logs.ToArray()
                Success = $false
            }
            return
        }

        # Step C: Get attendance records for each report
        foreach ($report in $reports) {
            $recordsUri = "$graphBase/v1.0/users/$meetingUserId/onlineMeetings/$($meeting.id)" +
                          "/attendanceReports/$($report.id)?`$expand=attendanceRecords"
            $detail = Invoke-GraphRest -Uri $recordsUri -Headers $headers

            foreach ($record in $detail.attendanceRecords) {
                $meetingSec = ($mEnd - $mStart).TotalSeconds
                $attPct     = if ($meetingSec -gt 0) {
                                  ($record.totalAttendanceInSeconds / $meetingSec) * 100
                              } else { 100 }

                $joinDt  = if ($record.attendanceIntervals -and @($record.attendanceIntervals).Count -gt 0) {
                               [datetime]$record.attendanceIntervals[0].joinDateTime
                           } else { $null }
                $leaveDt = if ($record.attendanceIntervals -and @($record.attendanceIntervals).Count -gt 0) {
                               [datetime]$record.attendanceIntervals[-1].leaveDateTime
                           } else { $null }

                $lateCutoff = $mStart.AddMinutes($lateThreshMin)
                $status = if ($attPct -lt $partialThreshPct)              { "Partial" }
                          elseif ($joinDt -and $joinDt -gt $lateCutoff)   { "Late"    }
                          else                                              { "Present" }

                $results.Add([PSCustomObject]@{
                    Date             = $mStart.ToString('yyyy-MM-dd')
                    TeacherName      = $wi.TeacherName
                    TeacherEmail     = $wi.TeacherEmail
                    Department       = $wi.TeacherDept
                    MeetingSubject   = $meeting.subject
                    MeetingStart     = $mStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
                    MeetingEnd       = $mEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
                    AttendeeName     = $record.identity.displayName
                    AttendeeEmail    = $record.emailAddress
                    JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ')  } else { '' }
                    LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                    DurationMinutes  = [math]::Round($record.totalAttendanceInSeconds / 60, 1)
                    AttendanceStatus = $status
                })
            }
        }

        $logs.Add("INFO | $($meeting.subject) (teacher=$($wi.TeacherEmail)) — $($reports.Count) report(s), $($results.Count) record(s)")
        $success = $true
    }
    catch {
        $logs.Add("ERROR | Failed processing meeting for $($wi.TeacherEmail): $_")
        $progress.Errors++
    }
    finally {
        # ── Update progress counter ──
        $progress.Done++
        $done = $progress.Done

        # Live progress bar (updates in real-time in the console)
        $pct = [math]::Min(100, [int](($done / [math]::Max($total, 1)) * 100))
        Write-Progress -Activity "Phase 2: Fetching attendance" `
            -Status "$done / $total meetings ($pct%) — errors: $($progress.Errors), throttled: $($progress.Throttled)" `
            -PercentComplete $pct

        # Periodic text log line (visible in console + scrollback)
        if ($done % $logEvery -eq 0 -or $done -eq $total) {
            $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            Write-Host "$ts | INFO |   Phase 2 progress: $done / $total ($pct%) — errors: $($progress.Errors), throttled: $($progress.Throttled)"
        }
    }

    # Return structured output from this worker
    [PSCustomObject]@{
        Results = $results.ToArray()
        Logs    = $logs.ToArray()
        Success = $success
    }
}

# Clear the progress bar
Write-Progress -Activity "Phase 2: Fetching attendance" -Completed

$phase2Timer.Stop()

# ── Collect results from parallel workers ──
$allResults      = [System.Collections.Generic.List[object]]::new()
$meetingsSuccess = 0
$meetingsFailed  = 0

foreach ($output in $parallelOutput) {
    # Replay log messages from the worker
    foreach ($logMsg in $output.Logs) {
        $parts = $logMsg -split ' \| ', 2
        $lvl   = if ($parts.Count -eq 2) { $parts[0].Trim() } else { "INFO" }
        $msg   = if ($parts.Count -eq 2) { $parts[1] }        else { $logMsg }
        Write-Log -Message "  [worker] $msg" -Level $lvl -LogsDir $config.logsDir
    }

    if ($output.Success) {
        $meetingsSuccess++
    }
    else {
        $meetingsFailed++
    }

    if ($output.Results -and $output.Results.Count -gt 0) {
        $allResults.AddRange($output.Results)
    }
}

Write-Log -Message "  Phase 2 completed in $([math]::Round($phase2Timer.Elapsed.TotalSeconds, 1))s — $($sharedProgress.Done) processed, $($sharedProgress.Errors) errors, $($sharedProgress.Throttled) throttle retries" -LogsDir $config.logsDir

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 3: Export results
# ══════════════════════════════════════════════════════════════════════════════

Write-Log -Message "Phase 3: Export" -LogsDir $config.logsDir
Write-Log -Message "  Meetings processed: $meetingsSuccess success, $meetingsFailed failed/skipped" -LogsDir $config.logsDir

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
