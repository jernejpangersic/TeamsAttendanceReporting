<#
.SYNOPSIS
    Extracts Teams meeting attendance data using Call Records as the discovery layer.
    Optimized v2: Oid-first filtering, Graph $batch API, and parallel attendance fetching.

.DESCRIPTION
    Instead of polling every teacher's calendar (35K calendarView calls), this script:
      1. Lists all Call Records for the target date range (1 paginated tenant-wide query)
      2. Filters to group calls involving teachers via Oid-first extraction from joinWebUrl
      3. Resolves meetings via batched Graph API calls (up to 20 per HTTP request)
      4. Pulls attendance reports and records in parallel (ForEach-Object -Parallel)
      5. Exports to Excel

    Optimizations over Get-AttendanceViaCallRecords.ps1:
      - Oid-first filtering: Extracts meeting creator from joinWebUrl before making API calls,
        eliminating hundreds of participant-lookup calls for non-teacher meetings
      - Graph $batch API: Groups up to 20 meeting resolution calls per HTTP request,
        reducing HTTP round-trip overhead by up to 20x
      - Parallel processing: Fetches attendance data for multiple meetings concurrently
        using ForEach-Object -Parallel (ThrottleLimit 10)
      - 429/503/504 retry: Automatic exponential backoff on throttled or transient errors

    Requires the additional permission: CallRecords.Read.All (application).

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.EXAMPLE
    .\Get-AttendanceViaCallRecords-v2.ps1
    .\Get-AttendanceViaCallRecords-v2.ps1 -TargetDate "2026-03-01"
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
    $logFile   = Join-Path $LogsDir "callrecords_v2_$(Get-Date -Format 'yyyy-MM-dd').log"

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

    $filePath = Join-Path $OutputDir "callrecords_v2_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

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
    Get-ChildItem -Path $OutputDir -Filter "callrecords_v2_*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            Remove-Item $_.FullName
            Write-Log -Message "Deleted old file: $($_.Name)" -LogsDir $LogsDir
        }
}

# ── Graph REST helper with retry for 429/503/504 ──
function Invoke-GraphRest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [string]$AccessToken,
        [object]$Body,
        [int]$MaxRetries = 3
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    if ($Body) { $headers['Content-Type'] = 'application/json' }

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        try {
            $params = @{
                Uri     = $Uri
                Method  = $Method
                Headers = $headers
            }
            if ($Body) {
                $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 }
            }
            return (Invoke-RestMethod @params)
        }
        catch {
            $statusCode = 0
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $isRetryable = ($statusCode -in @(429, 503, 504)) -or
                           ($null -eq $_.Exception.Response)  # network error

            if ($isRetryable -and $attempt -le $MaxRetries) {
                $retryAfter = $attempt * 5  # 5, 10, 15 seconds
                if ($statusCode -eq 429) {
                    try {
                        $retryValues = $_.Exception.Response.Headers.GetValues('Retry-After')
                        if ($retryValues) { $retryAfter = [math]::Max([int]$retryValues[0], $attempt * 2) }
                    }
                    catch { }
                }
                Start-Sleep -Seconds $retryAfter
            }
            else {
                throw
            }
        }
    }
}

# ── Graph $batch helper — sends up to 20 requests per HTTP call ──
function Invoke-GraphBatch {
    param(
        [array]$Requests,       # Array of @{ Id; Method; Url }
        [string]$AccessToken,
        [int]$BatchSize = 20
    )

    $allResponses = [System.Collections.Generic.Dictionary[string, object]]::new()

    for ($i = 0; $i -lt $Requests.Count; $i += $BatchSize) {
        $end   = [Math]::Min($i + $BatchSize - 1, $Requests.Count - 1)
        $chunk = $Requests[$i..$end]

        $batchBody = @{
            requests = @($chunk | ForEach-Object {
                @{
                    id     = $_.Id
                    method = $_.Method
                    url    = $_.Url
                }
            })
        }

        $resp = Invoke-GraphRest -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                                 -Method POST -AccessToken $AccessToken -Body $batchBody

        if ($resp.responses) {
            foreach ($r in $resp.responses) {
                $allResponses[$r.id] = $r
            }
        }
    }

    return $allResponses
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

# ── Script-level timer ──
$scriptTimer = [System.Diagnostics.Stopwatch]::StartNew()

Write-Log -Message "=== Call Records approach (v2 — optimized) ===" -LogsDir $config.logsDir
Write-Log -Message "Target date: $($TargetDate.ToString('yyyy-MM-dd')) | UTC range: $startIso to $endIso | Teachers: $($teachers.Count)" -LogsDir $config.logsDir

# ── Pre-capture config scalars ──
$lateThreshMin    = $config.lateThresholdMinutes
$partialThreshPct = $config.partialThresholdPercent

# ── Connect to Graph (for Phase 1 paginated listing) ──
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
Write-Log -Message "Connected to Microsoft Graph" -LogsDir $config.logsDir

# ── Acquire raw access token for $batch + parallel calls ──
# ForEach-Object -Parallel runspaces don't share the Graph SDK context,
# so we acquire a token via client credentials and use Invoke-RestMethod directly.
$tokenEndpoint = "https://login.microsoftonline.com/$($config.tenantId)/oauth2/v2.0/token"
$tokenBody = @{
    client_id     = $config.clientId
    client_secret = $config.clientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}
$tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $tokenBody `
                                   -ContentType "application/x-www-form-urlencoded"
$accessToken = $tokenResponse.access_token
Write-Log -Message "Acquired access token for batch/parallel operations" -LogsDir $config.logsDir

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

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 1: Discover meetings via Call Records
# ══════════════════════════════════════════════════════════════════════════════

Write-Progress -Activity "Teams Attendance Export" -Status "Phase 1: Discovering meetings via Call Records..." -PercentComplete 5
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

# ── Bucket 1: Meetings organized by teachers (direct organizer match) ──
$teacherMeetings = @($meetingRecords | Where-Object {
    $orgId = $_.organizer.user.id
    $orgId -and $teacherIdSet.Contains($orgId)
})

Write-Log -Message "  Organized by teachers in our list: $($teacherMeetings.Count)" -LogsDir $config.logsDir

# ── Bucket 2 (OPTIMIZED): Non-teacher-organized — Oid-first filtering ──
# Instead of making an API call per record to fetch participants, we first extract
# the meeting creator's Oid from the joinWebUrl (free string operation). If the Oid
# is not a teacher, we skip the record entirely — no API call needed.
$nonTeacherRecords = @($meetingRecords | Where-Object {
    $orgId = $_.organizer.user.id
    -not ($orgId -and $teacherIdSet.Contains($orgId))
})

$channelMeetings      = [System.Collections.Generic.List[object]]::new()
$needParticipantCheck = [System.Collections.Generic.List[object]]::new()
$skippedByOid         = 0
$addedByOid           = 0
$oidFilterIndex       = 0

Write-Progress -Activity "Teams Attendance Export" -Status "Phase 1: Oid-first filtering ($($nonTeacherRecords.Count) records)..." -PercentComplete 15

foreach ($cr in $nonTeacherRecords) {
    $oidFilterIndex++
    if ($oidFilterIndex % 100 -eq 0 -or $oidFilterIndex -eq $nonTeacherRecords.Count) {
        $oidPct = [math]::Min(100, [int](($oidFilterIndex / [math]::Max($nonTeacherRecords.Count, 1)) * 100))
        Write-Progress -Id 1 -Activity "Oid filtering" -Status "$oidFilterIndex / $($nonTeacherRecords.Count) ($oidPct%)" -PercentComplete $oidPct
    }
    # Extract meeting creator Oid from joinWebUrl FIRST (free — no API call)
    # Format: ...context=%7b%22Tid%22%3a%22{tenantId}%22%2c%22Oid%22%3a%22{userId}%22%7d
    $oidMatch = [regex]::Match($cr.joinWebUrl, 'Oid%22%3a%22([0-9a-f-]+)%22')

    if ($oidMatch.Success) {
        $creatorId = $oidMatch.Groups[1].Value
        if ($teacherIdSet.Contains($creatorId)) {
            # Creator IS a teacher — add directly, no API call needed
            $fakeParticipant = [PSCustomObject]@{ user = [PSCustomObject]@{ id = $creatorId } }
            $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue @($fakeParticipant) -Force
            $channelMeetings.Add($cr)
            $addedByOid++
        }
        else {
            # Creator is NOT a teacher — skip entirely (saves an API call)
            $skippedByOid++
        }
    }
    else {
        # No Oid in URL — need to fetch participants via API
        $needParticipantCheck.Add($cr)
    }
}

Write-Progress -Id 1 -Activity "Oid filtering" -Completed
Write-Log -Message "  Oid-first: $addedByOid added directly, $skippedByOid skipped (not teachers), $($needParticipantCheck.Count) need participant check" -LogsDir $config.logsDir

# Batch-fetch participants for records where Oid wasn't available (rare edge case)
$batchSkipped = 0
if ($needParticipantCheck.Count -gt 0) {
    Write-Log -Message "  Batch-fetching participants for $($needParticipantCheck.Count) record(s) without Oid..." -LogsDir $config.logsDir

    $participantBatchRequests = @($needParticipantCheck | ForEach-Object {
        @{
            Id     = $_.id
            Method = "GET"
            Url    = "/communications/callRecords/$($_.id)?`$select=id,participants"
        }
    })

    $participantBatchResponses = Invoke-GraphBatch -Requests $participantBatchRequests -AccessToken $accessToken

    foreach ($cr in $needParticipantCheck) {
        $resp = $participantBatchResponses[$cr.id]
        if (-not $resp -or $resp.status -ne 200) {
            Write-Log -Message "    Batch fetch failed for $($cr.id): status=$($resp.status)" -Level "WARN" -LogsDir $config.logsDir
            $batchSkipped++
            continue
        }

        $full = $resp.body
        $teacherParticipants = @($full.participants | Where-Object {
            $_.user.id -and $teacherIdSet.Contains($_.user.id)
        })

        if ($teacherParticipants.Count -gt 0) {
            $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue $teacherParticipants -Force
            $channelMeetings.Add($cr)
        }
        else {
            $batchSkipped++
        }
    }
}

Write-Log -Message "  Channel/external meetings with teacher involvement: $($channelMeetings.Count)" -LogsDir $config.logsDir
$totalSkipped = $skippedByOid + $batchSkipped
if ($totalSkipped -gt 0) {
    Write-Log -Message "  Total skipped (no teacher involvement): $totalSkipped" -LogsDir $config.logsDir
}

# ── Pre-deduplicate by joinWebUrl before Phase 2 ──
$allMeetingsToProcess = [System.Collections.Generic.List[object]]::new()
$seenUrls = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

foreach ($cr in (@($teacherMeetings) + @($channelMeetings))) {
    if ($cr.joinWebUrl -and $seenUrls.Add($cr.joinWebUrl)) {
        $allMeetingsToProcess.Add($cr)
    }
}

$dupCount = ($teacherMeetings.Count + $channelMeetings.Count) - $allMeetingsToProcess.Count
Write-Log -Message "  Unique meetings to process: $($allMeetingsToProcess.Count) (removed $dupCount duplicates)" -LogsDir $config.logsDir

if ($allMeetingsToProcess.Count -eq 0) {
    Write-Log -Message "No meetings to process — skipping Phase 2" -Level "WARN" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2: Resolve meetings via $batch, then fetch attendance in parallel
# ══════════════════════════════════════════════════════════════════════════════

$resolvedMeetings = [System.Collections.Generic.List[object]]::new()

if ($allMeetingsToProcess.Count -gt 0) {
    Write-Progress -Activity "Teams Attendance Export" -Status "Phase 2a: Resolving meetings via `$batch API..." -PercentComplete 35
    Write-Log -Message "Phase 2a: Resolving $($allMeetingsToProcess.Count) meetings via `$batch API..." -LogsDir $config.logsDir

    # ── Step A: Build candidate info for each meeting ──
    $meetingCandidates = [System.Collections.Generic.List[object]]::new()
    foreach ($cr in $allMeetingsToProcess) {
        $candidateUserIds = [System.Collections.Generic.List[string]]::new()
        $teacherForRow    = $null

        $orgId = $cr.organizer.user.id
        if ($orgId -and $teacherIdSet.Contains($orgId)) {
            $candidateUserIds.Add($orgId)
            $teacherForRow = $teacherLookup[$orgId]
        }
        else {
            if ($orgId) { $candidateUserIds.Add($orgId) }
            if ($cr.'_teacherParticipants') {
                foreach ($p in $cr._teacherParticipants) {
                    $pid2 = $p.user.id
                    if ($pid2 -and $teacherIdSet.Contains($pid2) -and -not $candidateUserIds.Contains($pid2)) {
                        $candidateUserIds.Add($pid2)
                        if (-not $teacherForRow) { $teacherForRow = $teacherLookup[$pid2] }
                    }
                }
            }
        }

        if ($teacherForRow -and $candidateUserIds.Count -gt 0) {
            $meetingCandidates.Add([PSCustomObject]@{
                CallRecord       = $cr
                CandidateUserIds = $candidateUserIds
                TeacherForRow    = $teacherForRow
                EncodedJoinUrl   = [Uri]::EscapeDataString($cr.joinWebUrl)
            })
        }
        else {
            Write-Log -Message "  No teacher candidate for call record $($cr.id) — skipping" -Level "WARN" -LogsDir $config.logsDir
        }
    }

    Write-Log -Message "  Meeting candidates prepared: $($meetingCandidates.Count)" -LogsDir $config.logsDir

    # ── Batch resolve meetings — try first candidate for each ──
    $batchRequests = @($meetingCandidates | ForEach-Object {
        $uid = $_.CandidateUserIds[0]
        @{
            Id     = $_.CallRecord.id
            Method = "GET"
            Url    = "/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$($_.EncodedJoinUrl)'"
        }
    })

    $batchResponses = if ($batchRequests.Count -gt 0) {
        Invoke-GraphBatch -Requests $batchRequests -AccessToken $accessToken
    }
    else {
        [System.Collections.Generic.Dictionary[string, object]]::new()
    }

    # Parse batch results
    $retryList = [System.Collections.Generic.List[object]]::new()

    foreach ($mc in $meetingCandidates) {
        $resp = $batchResponses[$mc.CallRecord.id]
        $meeting = $null

        if ($resp -and $resp.status -eq 200 -and $resp.body.value) {
            $meetingValues = @($resp.body.value)
            if ($meetingValues.Count -gt 0) {
                $meeting = $meetingValues[0]
            }
        }

        if ($meeting) {
            $resolvedMeetings.Add([PSCustomObject]@{
                Meeting       = $meeting
                MeetingUserId = $mc.CandidateUserIds[0]
                TeacherForRow = $mc.TeacherForRow
                CallRecord    = $mc.CallRecord
            })
        }
        elseif ($mc.CandidateUserIds.Count -gt 1) {
            # First candidate failed — queue for retry
            $retryList.Add($mc)
        }
        else {
            Write-Log -Message "  No meeting found for call record $($mc.CallRecord.id) — subject unknown" -Level "WARN" -LogsDir $config.logsDir
        }
    }

    # ── Retry failed resolutions with alternate candidates (via second batch) ──
    if ($retryList.Count -gt 0) {
        Write-Log -Message "  Retrying $($retryList.Count) meeting(s) with alternate candidates..." -LogsDir $config.logsDir

        # Build batch requests trying the second candidate
        $retryBatchRequests = @($retryList | ForEach-Object {
            $uid = $_.CandidateUserIds[1]
            @{
                Id     = $_.CallRecord.id
                Method = "GET"
                Url    = "/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$($_.EncodedJoinUrl)'"
            }
        })

        $retryBatchResponses = Invoke-GraphBatch -Requests $retryBatchRequests -AccessToken $accessToken

        foreach ($mc in $retryList) {
            $resp = $retryBatchResponses[$mc.CallRecord.id]
            $meeting = $null

            if ($resp -and $resp.status -eq 200 -and $resp.body.value) {
                $meetingValues = @($resp.body.value)
                if ($meetingValues.Count -gt 0) { $meeting = $meetingValues[0] }
            }

            if ($meeting) {
                $resolvedMeetings.Add([PSCustomObject]@{
                    Meeting       = $meeting
                    MeetingUserId = $mc.CandidateUserIds[1]
                    TeacherForRow = $mc.TeacherForRow
                    CallRecord    = $mc.CallRecord
                })
            }
            else {
                # Try remaining candidates sequentially (rare — 3+ candidates)
                $found = $false
                for ($ci = 2; $ci -lt $mc.CandidateUserIds.Count; $ci++) {
                    $uid = $mc.CandidateUserIds[$ci]
                    try {
                        $meetingUri = "https://graph.microsoft.com/v1.0/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$($mc.EncodedJoinUrl)'"
                        $result = Invoke-GraphRest -Uri $meetingUri -AccessToken $accessToken
                        $m = $result.value | Select-Object -First 1
                        if ($m) {
                            $resolvedMeetings.Add([PSCustomObject]@{
                                Meeting       = $m
                                MeetingUserId = $uid
                                TeacherForRow = $mc.TeacherForRow
                                CallRecord    = $mc.CallRecord
                            })
                            $found = $true
                            break
                        }
                    }
                    catch { }
                }
                if (-not $found) {
                    Write-Log -Message "  No meeting found for call record $($mc.CallRecord.id) after $($mc.CandidateUserIds.Count) candidates" -Level "WARN" -LogsDir $config.logsDir
                }
            }
        }
    }

    Write-Log -Message "  Resolved meetings: $($resolvedMeetings.Count) of $($meetingCandidates.Count)" -LogsDir $config.logsDir

    # ── Step B+C: Fetch attendance in parallel ──
    Write-Progress -Activity "Teams Attendance Export" -Status "Phase 2b: Fetching attendance for $($resolvedMeetings.Count) meetings..." -PercentComplete 50
    Write-Log -Message "Phase 2b: Fetching attendance for $($resolvedMeetings.Count) meetings in parallel (ThrottleLimit 10)..." -LogsDir $config.logsDir
}

$allResults = @()

if ($resolvedMeetings.Count -gt 0) {
    # ── Thread-safe progress state shared across parallel workers ──
    $totalMeetings  = $resolvedMeetings.Count
    $sharedProgress = [hashtable]::Synchronized(@{ Done = 0; Errors = 0; Throttled = 0 })
    # Log a text line every ~10% of total, but at least every 50 and at most every 500
    $logInterval    = [math]::Max(50, [math]::Min(500, [math]::Ceiling($totalMeetings / 10)))

    $phase2Timer = [System.Diagnostics.Stopwatch]::StartNew()

    $allResults = $resolvedMeetings | ForEach-Object -Parallel {
        $rm             = $_
        $token          = $using:accessToken
        $lateMin        = $using:lateThreshMin
        $partialPct     = $using:partialThreshPct
        $graphBase      = "https://graph.microsoft.com/v1.0"
        $headers        = @{ Authorization = "Bearer $token" }

        # Import shared progress state
        $progress       = $using:sharedProgress
        $total          = $using:totalMeetings
        $logEvery       = $using:logInterval

        $meeting        = $rm.Meeting
        $meetingUserId  = $rm.MeetingUserId
        $teacher        = $rm.TeacherForRow
        $cr             = $rm.CallRecord

        # Use actual times from call record (not scheduled times)
        $mStart = [datetime]$cr.startDateTime
        $mEnd   = [datetime]$cr.endDateTime

        # Local retry helper (functions from parent scope aren't available in parallel runspaces)
        $invokeWithRetry = {
            param([string]$Uri, [hashtable]$Hdrs, [int]$Retries = 5)
            for ($att = 1; $att -le ($Retries + 1); $att++) {
                try {
                    return (Invoke-RestMethod -Uri $Uri -Headers $Hdrs)
                }
                catch {
                    $sc = 0
                    if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
                    $retryable = ($sc -in @(429, 503, 504)) -or ($null -eq $_.Exception.Response)
                    if ($retryable -and $att -le $Retries) {
                        if ($sc -eq 429) { $progress.Throttled++ }
                        $waitSec = $att * 5  # 5, 10, 15, 20, 25s
                        if ($sc -eq 429) {
                            try {
                                $raValues = $_.Exception.Response.Headers.GetValues('Retry-After')
                                if ($raValues) { $waitSec = [math]::Max([int]$raValues[0], $att * 2) }
                            } catch { }
                        }
                        Start-Sleep -Seconds $waitSec
                    }
                    else { throw }
                }
            }
        }

        try {
            # Get attendance reports (with pagination)
            $reportsUri  = "$graphBase/users/$meetingUserId/onlineMeetings/$($meeting.id)/attendanceReports"
            $reportsResp = & $invokeWithRetry $reportsUri $headers
            $reports     = [System.Collections.Generic.List[object]]::new()
            if ($reportsResp.value) { $reports.AddRange(@($reportsResp.value)) }

            $nextLink = $reportsResp.'@odata.nextLink'
            while ($nextLink) {
                $page = & $invokeWithRetry $nextLink $headers
                if ($page.value) { $reports.AddRange(@($page.value)) }
                $nextLink = $page.'@odata.nextLink'
            }

            if ($reports.Count -eq 0) { return }

            foreach ($report in $reports) {
                # Get attendance records (with $expand)
                $recordsUri = "$graphBase/users/$meetingUserId/onlineMeetings/$($meeting.id)" +
                              "/attendanceReports/$($report.id)?`$expand=attendanceRecords"
                $detail = & $invokeWithRetry $recordsUri $headers

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

                    $lateCutoff = $mStart.AddMinutes($lateMin)
                    $status = if ($attPct -lt $partialPct)                { "Partial" }
                              elseif ($joinDt -and $joinDt -gt $lateCutoff) { "Late"    }
                              else                                          { "Present" }

                    # Emit row — collected by ForEach-Object -Parallel
                    [PSCustomObject]@{
                        Date             = $mStart.ToString('yyyy-MM-dd')
                        TeacherName      = $teacher.displayName
                        TeacherEmail     = $teacher.mail
                        Department       = $teacher.department
                        MeetingSubject   = $meeting.subject
                        MeetingStart     = $mStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
                        MeetingEnd       = $mEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
                        AttendeeName     = $record.identity.displayName
                        AttendeeEmail    = $record.emailAddress
                        JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ')  } else { '' }
                        LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                        DurationMinutes  = [math]::Round($record.totalAttendanceInSeconds / 60, 1)
                        AttendanceStatus = $status
                    }
                }
            }
        }
        catch {
            $progress.Errors++
            Write-Warning "Failed processing meeting '$($meeting.subject)' for $($teacher.mail): $_"
        }
        finally {
            # ── Update progress counter ──
            $progress.Done++
            $done = $progress.Done

            # Live progress bar (updates in real-time in the console)
            $pct = [math]::Min(100, [int](($done / [math]::Max($total, 1)) * 100))
            Write-Progress -Activity "Phase 2b: Fetching attendance" `
                -Status "$done / $total meetings ($pct%) — errors: $($progress.Errors), throttled: $($progress.Throttled)" `
                -PercentComplete $pct

            # Periodic text log line (visible in console + scrollback)
            if ($done % $logEvery -eq 0 -or $done -eq $total) {
                $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                Write-Host "$ts | INFO |   Phase 2b progress: $done / $total ($pct%) — errors: $($progress.Errors), throttled: $($progress.Throttled)"
            }
        }
    } -ThrottleLimit 10

    # Clear the progress bar
    Write-Progress -Activity "Phase 2b: Fetching attendance" -Completed

    $phase2Timer.Stop()
    $allResults = @($allResults)  # Force to array (parallel output is a stream)

    Write-Log -Message "  Phase 2b completed in $([math]::Round($phase2Timer.Elapsed.TotalSeconds, 1))s — $($sharedProgress.Done) processed, $($sharedProgress.Errors) errors, $($sharedProgress.Throttled) throttle retries" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 3: Export results
# ══════════════════════════════════════════════════════════════════════════════

Write-Progress -Activity "Teams Attendance Export" -Status "Phase 3: Exporting to Excel..." -PercentComplete 90
Write-Log -Message "Phase 3: Export" -LogsDir $config.logsDir
Write-Log -Message "  Meetings resolved: $($resolvedMeetings.Count)" -LogsDir $config.logsDir
Write-Log -Message "  Attendance records: $($allResults.Count)" -LogsDir $config.logsDir

if ($allResults -and $allResults.Count -gt 0) {
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

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Write-Progress -Activity "Teams Attendance Export" -Completed

$scriptTimer.Stop()
Write-Log -Message "Total runtime: $([math]::Round($scriptTimer.Elapsed.TotalSeconds, 1))s" -LogsDir $config.logsDir

#endregion
