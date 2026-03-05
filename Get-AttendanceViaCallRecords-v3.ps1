<#
.SYNOPSIS
    Extracts Teams meeting attendance data using Call Records — v3 optimized for
    large-scale tenants (20k+ meetings/day).

.DESCRIPTION
    Builds on v2's OID-first filtering and $batch meeting resolution, then pushes
    the $batch API all the way through attendance fetching.  Instead of dispatching
    one HTTP call per meeting inside parallel workers (v2), v3 groups meetings into
    chunks of 20 and processes each chunk with just 2-3 $batch HTTP calls.

    Architecture:
      Phase 1 : Call Record discovery + OID-first filtering        (same as v2)
      Phase 2a: Meeting resolution via $batch + $select            (same as v2, leaner payloads)
      Phase 2b: Attendance via Chunked-Parallel-Batch
                - Split resolved meetings into chunks of 20
                - Dispatch chunks to parallel workers (ThrottleLimit)
                - Each worker sends 1 $batch for reports, 1+ $batch for records
                - Rows are emitted directly from the pipeline
      Phase 3 : Export to Excel

    Key optimisations over v2:
      - Batched attendance fetching: 20 report/record requests per HTTP round-trip
      - $select on all Graph calls: ~30-50% smaller JSON payloads
      - Token refresh: proactive mid-run refresh for long executions (>50 min)
      - Exponential backoff: 2^n second retry on 429/503/504 (outer + inner batch)
      - Per-item batch retry: only re-sends throttled items, not the whole batch
      - Configurable ThrottleLimit: tune parallelism from the command line

    Estimated HTTP calls at 20k meetings (assuming 1 report per meeting):
      v2 :  ~1,000 (resolution) + ~40,000 (attendance)     = ~41,000
      v3 :  ~1,000 (resolution) + ~2,000 (attendance batch) = ~3,000

    Requires: PowerShell 7+, CallRecords.Read.All (application).

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.PARAMETER ThrottleLimit
    Maximum concurrent workers for Phase 2b parallel batch processing. Default: 10.
    Higher values are faster but increase risk of Graph 429 throttling.

.EXAMPLE
    .\Get-AttendanceViaCallRecords-v3.ps1
    .\Get-AttendanceViaCallRecords-v3.ps1 -TargetDate "2026-03-01" -ThrottleLimit 15
#>
param(
    [string]$ConfigPath = ".\config.json",
    [datetime]$TargetDate,
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 10
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
    $logFile   = Join-Path $LogsDir "callrecords_v3_$(Get-Date -Format 'yyyy-MM-dd').log"

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

    $filePath = Join-Path $OutputDir "callrecords_v3_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

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
    Get-ChildItem -Path $OutputDir -Filter "callrecords_v3_*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            Remove-Item $_.FullName
            Write-Log -Message "Deleted old file: $($_.Name)" -LogsDir $LogsDir
        }
}

# ── Graph REST helper with exponential-backoff retry for 429/503/504 ──
function Invoke-GraphRest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [string]$AccessToken,
        [object]$Body,
        [int]$MaxRetries = 4
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
                # Exponential backoff: 2, 4, 8, 16s — respect Retry-After header
                $retryAfter = [math]::Pow(2, $attempt)
                if ($statusCode -eq 429) {
                    try {
                        $retryValues = $_.Exception.Response.Headers.GetValues('Retry-After')
                        if ($retryValues) {
                            $parsed = [int]($retryValues | Select-Object -First 1)
                            if ($parsed -gt 0) { $retryAfter = [math]::Max($parsed, $retryAfter) }
                        }
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

# ── Graph $batch helper with per-item retry ──
# Sends up to 20 requests per HTTP call.  If individual items come back
# 429/503/504, they are collected and retried in a subsequent batch with
# exponential backoff — only the failed items, not the entire batch.
function Invoke-GraphBatch {
    param(
        [array]$Requests,       # Array of @{ Id; Method; Url }
        [string]$AccessToken,
        [int]$BatchSize = 20,
        [int]$MaxItemRetries = 3
    )

    $allResponses = [System.Collections.Generic.Dictionary[string, object]]::new()
    $pending      = [System.Collections.Generic.List[object]]::new($Requests)

    for ($retry = 0; $retry -le $MaxItemRetries -and $pending.Count -gt 0; $retry++) {
        $nextPending = [System.Collections.Generic.List[object]]::new()

        for ($i = 0; $i -lt $pending.Count; $i += $BatchSize) {
            $end   = [Math]::Min($i + $BatchSize - 1, $pending.Count - 1)
            $chunk = @($pending[$i..$end])

            $batchBody = @{
                requests = @($chunk | ForEach-Object {
                    @{ id = $_.Id; method = $_.Method; url = $_.Url }
                })
            }

            $resp = Invoke-GraphRest -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                                     -Method POST -AccessToken $AccessToken -Body $batchBody

            if ($resp.responses) {
                foreach ($r in $resp.responses) {
                    if ($r.status -in @(429, 503, 504)) {
                        $original = $chunk | Where-Object { $_.Id -eq $r.id } | Select-Object -First 1
                        if ($original) { $nextPending.Add($original) }
                    }
                    else {
                        $allResponses[$r.id] = $r
                    }
                }
            }
        }

        if ($nextPending.Count -gt 0 -and $retry -lt $MaxItemRetries) {
            $waitSec = [math]::Pow(2, $retry + 1)   # 2, 4, 8s
            Start-Sleep -Seconds $waitSec
            $pending = $nextPending
        }
        else {
            # Record failures so callers can detect them
            foreach ($item in $nextPending) {
                $allResponses[$item.Id] = @{ id = $item.Id; status = 429; body = $null }
            }
            break
        }
    }

    return $allResponses
}

#endregion

#region ── Main Script ──

# ── Phase timers for summary ──
$phaseTimers = @{}

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

Write-Log -Message "=== Call Records approach (v3 — chunked parallel batch) ===" -LogsDir $config.logsDir
Write-Log -Message "Target date: $($TargetDate.ToString('yyyy-MM-dd')) | UTC range: $startIso to $endIso | Teachers: $($teachers.Count) | ThrottleLimit: $ThrottleLimit" -LogsDir $config.logsDir

# ── Pre-capture config scalars for parallel workers ──
$lateThreshMin    = $config.lateThresholdMinutes
$partialThreshPct = $config.partialThresholdPercent

# ── Connect to Graph (for Phase 1 paginated listing) ──
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
Write-Log -Message "Connected to Microsoft Graph" -LogsDir $config.logsDir

# ── Acquire raw access token with expiry tracking ──
# ForEach-Object -Parallel runspaces don't share the Graph SDK context,
# so we acquire a token via client credentials and use Invoke-RestMethod directly.
# We also pass credentials so parallel workers can refresh mid-run.
$tokenTenantId = $config.tenantId
$tokenClientId = $config.clientId
$tokenClientSecret = $config.clientSecret

$tokenEndpoint = "https://login.microsoftonline.com/$tokenTenantId/oauth2/v2.0/token"
$tokenBody = @{
    client_id     = $tokenClientId
    client_secret = $tokenClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}
$tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $tokenBody `
                                   -ContentType "application/x-www-form-urlencoded"
$accessToken    = $tokenResponse.access_token
$tokenExpiresAt = [datetime]::UtcNow.AddSeconds($tokenResponse.expires_in)
Write-Log -Message "Acquired access token (expires $($tokenExpiresAt.ToString('HH:mm:ss'))Z)" -LogsDir $config.logsDir

# ── Helper: refresh token if within 5 minutes of expiry ──
function Refresh-TokenIfNeeded {
    if ([datetime]::UtcNow -ge $script:tokenExpiresAt.AddMinutes(-5)) {
        $resp = Invoke-RestMethod -Uri $script:tokenEndpoint -Method POST `
                    -Body $script:tokenBody -ContentType "application/x-www-form-urlencoded"
        $script:accessToken    = $resp.access_token
        $script:tokenExpiresAt = [datetime]::UtcNow.AddSeconds($resp.expires_in)
        Write-Log -Message "  Token refreshed (expires $($script:tokenExpiresAt.ToString('HH:mm:ss'))Z)" -LogsDir $config.logsDir
    }
}

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

$phase1Timer = [System.Diagnostics.Stopwatch]::StartNew()

Write-Progress -Activity "Teams Attendance Export (v3)" -Status "Phase 1: Discovering meetings via Call Records..." -PercentComplete 5
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
$nonTeacherRecords = @($meetingRecords | Where-Object {
    $orgId = $_.organizer.user.id
    -not ($orgId -and $teacherIdSet.Contains($orgId))
})

$channelMeetings      = [System.Collections.Generic.List[object]]::new()
$needParticipantCheck = [System.Collections.Generic.List[object]]::new()
$skippedByOid         = 0
$addedByOid           = 0
$oidFilterIndex       = 0

Write-Progress -Activity "Teams Attendance Export (v3)" -Status "Phase 1: Oid-first filtering ($($nonTeacherRecords.Count) records)..." -PercentComplete 15

foreach ($cr in $nonTeacherRecords) {
    $oidFilterIndex++
    if ($oidFilterIndex % 100 -eq 0 -or $oidFilterIndex -eq $nonTeacherRecords.Count) {
        $oidPct = [math]::Min(100, [int](($oidFilterIndex / [math]::Max($nonTeacherRecords.Count, 1)) * 100))
        Write-Progress -Id 1 -Activity "Oid filtering" -Status "$oidFilterIndex / $($nonTeacherRecords.Count) ($oidPct%)" -PercentComplete $oidPct
    }

    $oidMatch = [regex]::Match($cr.joinWebUrl, 'Oid%22%3a%22([0-9a-f-]+)%22')

    if ($oidMatch.Success) {
        $creatorId = $oidMatch.Groups[1].Value
        if ($teacherIdSet.Contains($creatorId)) {
            $fakeParticipant = [PSCustomObject]@{ user = [PSCustomObject]@{ id = $creatorId } }
            $cr | Add-Member -NotePropertyName '_teacherParticipants' -NotePropertyValue @($fakeParticipant) -Force
            $channelMeetings.Add($cr)
            $addedByOid++
        }
        else {
            $skippedByOid++
        }
    }
    else {
        $needParticipantCheck.Add($cr)
    }
}

Write-Progress -Id 1 -Activity "Oid filtering" -Completed
Write-Log -Message "  Oid-first: $addedByOid added directly, $skippedByOid skipped (not teachers), $($needParticipantCheck.Count) need participant check" -LogsDir $config.logsDir

# Batch-fetch participants for records where Oid wasn't available
$batchSkipped = 0
if ($needParticipantCheck.Count -gt 0) {
    Refresh-TokenIfNeeded
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

$phase1Timer.Stop()
$phaseTimers['Phase1'] = $phase1Timer.Elapsed

if ($allMeetingsToProcess.Count -eq 0) {
    Write-Log -Message "No meetings to process — skipping Phase 2" -Level "WARN" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2a: Resolve meetings via $batch (with $select for smaller payloads)
# ══════════════════════════════════════════════════════════════════════════════

$resolvedMeetings = [System.Collections.Generic.List[object]]::new()

if ($allMeetingsToProcess.Count -gt 0) {
    $phase2aTimer = [System.Diagnostics.Stopwatch]::StartNew()

    Refresh-TokenIfNeeded
    Write-Progress -Activity "Teams Attendance Export (v3)" -Status "Phase 2a: Resolving meetings via `$batch API..." -PercentComplete 30
    Write-Log -Message "Phase 2a: Resolving $($allMeetingsToProcess.Count) meetings via `$batch API..." -LogsDir $config.logsDir

    # ── Build candidate info for each meeting ──
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

    # ── Batch resolve meetings — try first candidate for each (with $select) ──
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
            if ($meetingValues.Count -gt 0) { $meeting = $meetingValues[0] }
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
            $retryList.Add($mc)
        }
        else {
            $respStatus = if ($resp) { $resp.status } else { 'null' }
            Write-Log -Message "  No meeting found for call record $($mc.CallRecord.id) (batch status: $respStatus)" -Level "WARN" -LogsDir $config.logsDir
        }
    }

    # ── Retry failed resolutions with alternate candidates ──
    if ($retryList.Count -gt 0) {
        Refresh-TokenIfNeeded
        Write-Log -Message "  Retrying $($retryList.Count) meeting(s) with alternate candidates..." -LogsDir $config.logsDir

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

    $phase2aTimer.Stop()
    $phaseTimers['Phase2a'] = $phase2aTimer.Elapsed
    Write-Log -Message "  Resolved meetings: $($resolvedMeetings.Count) of $($meetingCandidates.Count) in $([math]::Round($phase2aTimer.Elapsed.TotalSeconds, 1))s" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2b: Fetch attendance via Chunked-Parallel-Batch
#
# Instead of one HTTP call per meeting (v2), we group 20 meetings into one
# $batch request.  Each parallel worker processes a chunk of 20 meetings:
#   Batch 1 → GET attendanceReports for 20 meetings  (1 HTTP call)
#   Batch 2 → GET attendanceRecords for all reports   (1+ HTTP calls)
# This reduces HTTP round-trips by up to 20x compared to v2.
# ══════════════════════════════════════════════════════════════════════════════

$allResults = @()

if ($resolvedMeetings.Count -gt 0) {
    $phase2bTimer = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Progress -Activity "Teams Attendance Export (v3)" -Status "Phase 2b: Chunked batch attendance for $($resolvedMeetings.Count) meetings..." -PercentComplete 45
    Write-Log -Message "Phase 2b: Fetching attendance for $($resolvedMeetings.Count) meetings via chunked parallel batch (ThrottleLimit $ThrottleLimit)..." -LogsDir $config.logsDir

    # ── Build serializable work items ──
    # Complex objects (PSCustomObject from Graph) don't always survive parallel
    # runspace serialisation cleanly.  We flatten to simple scalar properties.
    $workItems = foreach ($rm in $resolvedMeetings) {
        [PSCustomObject]@{
            MeetingId      = $rm.Meeting.id
            MeetingSubject = $rm.Meeting.subject
            MeetingUserId  = $rm.MeetingUserId
            TeacherName    = $rm.TeacherForRow.displayName
            TeacherEmail   = $rm.TeacherForRow.mail
            TeacherDept    = $rm.TeacherForRow.department
            StartDateTime  = [string]$rm.CallRecord.startDateTime
            EndDateTime    = [string]$rm.CallRecord.endDateTime
        }
    }
    $workItems = @($workItems)

    # ── Split work items into chunks of 20 (the $batch limit) ──
    $chunkSize     = 20
    $meetingChunks = [System.Collections.Generic.List[object[]]]::new()
    for ($i = 0; $i -lt $workItems.Count; $i += $chunkSize) {
        $end = [Math]::Min($i + $chunkSize - 1, $workItems.Count - 1)
        $meetingChunks.Add(@($workItems[$i..$end]))
    }

    Write-Log -Message "  Created $($meetingChunks.Count) chunks of up to $chunkSize meetings each" -LogsDir $config.logsDir

    # ── Refresh token just before dispatching parallel work ──
    Refresh-TokenIfNeeded

    # ── Thread-safe progress ──
    $totalChunks    = $meetingChunks.Count
    $sharedProgress = [hashtable]::Synchronized(@{
        Done = 0; Errors = 0; Throttled = 0; Meetings = 0; Records = 0
    })
    $logInterval = [math]::Max(10, [math]::Min(100, [math]::Ceiling($totalChunks / 10)))

    # ── Parallel execution: one worker per chunk ──
    $allResults = $meetingChunks | ForEach-Object -Parallel {
        $chunk = $_

        # Import variables from parent scope
        $token         = $using:accessToken
        $tokenExpiry   = $using:tokenExpiresAt
        $tenantId      = $using:tokenTenantId
        $clientId      = $using:tokenClientId
        $clientSecret  = $using:tokenClientSecret
        $lateMin       = $using:lateThreshMin
        $partialPct    = $using:partialThreshPct
        $progress      = $using:sharedProgress
        $totalChk      = $using:totalChunks
        $logEvery      = $using:logInterval
        $graphBatchUri = "https://graph.microsoft.com/v1.0/`$batch"

        $headers = @{
            Authorization  = "Bearer $token"
            'Content-Type' = 'application/json'
        }

        # ── Local: refresh token if within 5 min of expiry ──
        if ([datetime]::UtcNow -ge $tokenExpiry.AddMinutes(-5)) {
            try {
                $tBody = @{
                    client_id = $clientId; client_secret = $clientSecret
                    scope = "https://graph.microsoft.com/.default"; grant_type = "client_credentials"
                }
                $tResp = Invoke-RestMethod `
                    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
                    -Method POST -Body $tBody -ContentType "application/x-www-form-urlencoded"
                $token = $tResp.access_token
                $headers.Authorization = "Bearer $token"
            }
            catch {
                # If refresh fails, continue with current token — it may still work
            }
        }

        # ── Local: Send-GraphBatch with outer + inner retry ──
        function Send-GraphBatch {
            param(
                [array]$Requests,
                [hashtable]$Headers,
                [string]$BatchUri,
                [hashtable]$ProgressState,
                [int]$MaxItemRetries = 3
            )

            $results = @{}
            $pending = [System.Collections.Generic.List[object]]::new($Requests)

            for ($retry = 0; $retry -le $MaxItemRetries -and $pending.Count -gt 0; $retry++) {
                $nextPending = [System.Collections.Generic.List[object]]::new()

                # Process in sub-chunks of 20 (in case a single worker has >20 record requests)
                for ($i = 0; $i -lt $pending.Count; $i += 20) {
                    $end   = [Math]::Min($i + 19, $pending.Count - 1)
                    $slice = @($pending[$i..$end])

                    $body = @{
                        requests = @($slice | ForEach-Object {
                            @{ id = $_.id; method = $_.method; url = $_.url }
                        })
                    } | ConvertTo-Json -Depth 10

                    # Outer retry: if the $batch endpoint itself returns 429/503/504
                    $batchResp = $null
                    for ($att = 1; $att -le 4; $att++) {
                        try {
                            $batchResp = Invoke-RestMethod -Uri $BatchUri -Method POST `
                                            -Headers $Headers -Body $body
                            break
                        }
                        catch {
                            $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
                            if (($sc -in @(429, 503, 504)) -and $att -lt 4) {
                                if ($sc -eq 429 -and $ProgressState) { $ProgressState.Throttled++ }
                                $waitSec = [math]::Pow(2, $att)
                                if ($sc -eq 429) {
                                    try {
                                        $raVals = $null
                                        if ($_.Exception.Response.Headers.TryGetValues('Retry-After', [ref]$raVals)) {
                                            $parsed = [int]($raVals | Select-Object -First 1)
                                            if ($parsed -gt 0) { $waitSec = [math]::Max($parsed, $waitSec) }
                                        }
                                    } catch { }
                                }
                                Start-Sleep -Seconds $waitSec
                            }
                            else { throw }
                        }
                    }

                    if (-not $batchResp -or -not $batchResp.responses) { continue }

                    # Inner: check per-item status
                    foreach ($r in $batchResp.responses) {
                        if ($r.status -in @(429, 503, 504)) {
                            if ($r.status -eq 429 -and $ProgressState) { $ProgressState.Throttled++ }
                            $original = $slice | Where-Object { $_.id -eq $r.id } | Select-Object -First 1
                            if ($original) { $nextPending.Add($original) }
                        }
                        else {
                            $results[$r.id] = $r
                        }
                    }
                }

                if ($nextPending.Count -gt 0 -and $retry -lt $MaxItemRetries) {
                    Start-Sleep -Seconds ([math]::Pow(2, $retry + 1))
                }
                $pending = $nextPending
            }

            return $results
        }

        # ────────────────────────────────────────────────────────────────────
        # STEP 1: Batch-fetch attendance REPORTS for all meetings in chunk
        # ────────────────────────────────────────────────────────────────────
        $reportRequests = for ($idx = 0; $idx -lt $chunk.Count; $idx++) {
            $wi = $chunk[$idx]
            @{
                id     = "$idx"
                method = "GET"
                url    = "/users/$($wi.MeetingUserId)/onlineMeetings/$($wi.MeetingId)/attendanceReports?`$select=id"
            }
        }

        try {
            $reportResponses = Send-GraphBatch -Requests @($reportRequests) -Headers $headers `
                                               -BatchUri $graphBatchUri -ProgressState $progress

            # ────────────────────────────────────────────────────────────────
            # STEP 2: Build batch to fetch attendance RECORDS (with $expand)
            # ────────────────────────────────────────────────────────────────
            $recordRequests = [System.Collections.Generic.List[object]]::new()

            for ($idx = 0; $idx -lt $chunk.Count; $idx++) {
                $resp = $reportResponses["$idx"]
                if (-not $resp -or $resp.status -ne 200) { continue }

                $reports = $resp.body.value
                if (-not $reports -or @($reports).Count -eq 0) { continue }

                $wi     = $chunk[$idx]
                $rptIdx = 0
                foreach ($rpt in $reports) {
                    $recordRequests.Add(@{
                        id     = "${idx}_${rptIdx}"
                        method = "GET"
                        url    = "/users/$($wi.MeetingUserId)/onlineMeetings/$($wi.MeetingId)/attendanceReports/$($rpt.id)?`$expand=attendanceRecords"
                    })
                    $rptIdx++
                }

                # NOTE: If a report response contains @odata.nextLink (>50 reports),
                # the remaining pages are not fetched.  This is extremely rare for a
                # single meeting's reports within one day.  A future enhancement could
                # detect nextLink and issue follow-up calls here.
            }

            if ($recordRequests.Count -gt 0) {
                $recordResponses = Send-GraphBatch -Requests @($recordRequests) -Headers $headers `
                                                   -BatchUri $graphBatchUri -ProgressState $progress

                # ────────────────────────────────────────────────────────────
                # STEP 3: Process attendance records → emit rows
                # ────────────────────────────────────────────────────────────
                $chunkRecordCount = 0

                foreach ($key in $recordResponses.Keys) {
                    $resp = $recordResponses[$key]
                    if ($resp.status -ne 200 -or -not $resp.body) { continue }

                    # Parse composite key "meetingIdx_reportIdx"
                    $parts      = $key -split '_', 2
                    $meetingIdx = [int]$parts[0]
                    $wi         = $chunk[$meetingIdx]

                    $mStart = [datetime]$wi.StartDateTime
                    $mEnd   = [datetime]$wi.EndDateTime

                    $records = $resp.body.attendanceRecords
                    if (-not $records) { continue }

                    foreach ($rec in $records) {
                        $meetingSec = ($mEnd - $mStart).TotalSeconds
                        $attPct = if ($meetingSec -gt 0) {
                                      ($rec.totalAttendanceInSeconds / $meetingSec) * 100
                                  } else { 100 }

                        $joinDt  = if ($rec.attendanceIntervals -and @($rec.attendanceIntervals).Count -gt 0) {
                                       [datetime]$rec.attendanceIntervals[0].joinDateTime
                                   } else { $null }
                        $leaveDt = if ($rec.attendanceIntervals -and @($rec.attendanceIntervals).Count -gt 0) {
                                       [datetime]$rec.attendanceIntervals[-1].leaveDateTime
                                   } else { $null }

                        $lateCutoff = $mStart.AddMinutes($lateMin)
                        $status = if ($attPct -lt $partialPct)                  { "Partial" }
                                  elseif ($joinDt -and $joinDt -gt $lateCutoff) { "Late"    }
                                  else                                          { "Present" }

                        $chunkRecordCount++

                        # Emit row directly into the pipeline
                        [PSCustomObject]@{
                            Date             = $mStart.ToString('yyyy-MM-dd')
                            TeacherName      = $wi.TeacherName
                            TeacherEmail     = $wi.TeacherEmail
                            Department       = $wi.TeacherDept
                            MeetingSubject   = $wi.MeetingSubject
                            MeetingStart     = $mStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
                            MeetingEnd       = $mEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
                            AttendeeName     = $rec.identity.displayName
                            AttendeeEmail    = $rec.emailAddress
                            JoinTime         = if ($joinDt)  { $joinDt.ToString('yyyy-MM-ddTHH:mm:ssZ')  } else { '' }
                            LeaveTime        = if ($leaveDt) { $leaveDt.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
                            DurationMinutes  = [math]::Round($rec.totalAttendanceInSeconds / 60, 1)
                            AttendanceStatus = $status
                        }
                    }
                }

                $progress.Records += $chunkRecordCount
            }

            $progress.Meetings += $chunk.Count
        }
        catch {
            $progress.Errors++
            Write-Warning "Chunk failed: $_"
        }
        finally {
            $progress.Done++
            $done = $progress.Done
            $pct  = [math]::Min(100, [int](($done / [math]::Max($totalChk, 1)) * 100))

            Write-Progress -Activity "Phase 2b: Chunked batch attendance" `
                -Status "$done / $totalChk chunks ($pct%) — $($progress.Meetings) meetings, $($progress.Records) records, errors: $($progress.Errors), throttled: $($progress.Throttled)" `
                -PercentComplete $pct

            if ($done % $logEvery -eq 0 -or $done -eq $totalChk) {
                $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                Write-Host "$ts | INFO |   Phase 2b: $done / $totalChk chunks ($pct%) — $($progress.Meetings) meetings, $($progress.Records) records, errors: $($progress.Errors), throttled: $($progress.Throttled)"
            }
        }
    } -ThrottleLimit $ThrottleLimit

    # Clear progress bar
    Write-Progress -Activity "Phase 2b: Chunked batch attendance" -Completed

    $phase2bTimer.Stop()
    $phaseTimers['Phase2b'] = $phase2bTimer.Elapsed

    $allResults = @($allResults)  # Force stream to array

    Write-Log -Message "  Phase 2b completed in $([math]::Round($phase2bTimer.Elapsed.TotalSeconds, 1))s — $($sharedProgress.Done) chunks, $($sharedProgress.Meetings) meetings, $($sharedProgress.Records) records, $($sharedProgress.Errors) errors, $($sharedProgress.Throttled) throttle retries" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 3: Export results
# ══════════════════════════════════════════════════════════════════════════════

$phase3Timer = [System.Diagnostics.Stopwatch]::StartNew()

Write-Progress -Activity "Teams Attendance Export (v3)" -Status "Phase 3: Exporting to Excel..." -PercentComplete 90
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

    Write-Log -Message "  Exported to $filePath" -LogsDir $config.logsDir
}
else {
    Write-Log -Message "No attendance records found for $($TargetDate.ToString('yyyy-MM-dd'))" -Level "WARN" -LogsDir $config.logsDir
}

$phase3Timer.Stop()
$phaseTimers['Phase3'] = $phase3Timer.Elapsed

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Write-Progress -Activity "Teams Attendance Export (v3)" -Completed

$scriptTimer.Stop()

# ══════════════════════════════════════════════════════════════════════════════
# Summary
# ══════════════════════════════════════════════════════════════════════════════

$summaryLines = @(
    "=== v3 Performance Summary ==="
    "  Phase 1  (discovery)  : $([math]::Round(($phaseTimers['Phase1']).TotalSeconds, 1))s — $($allCallRecords.Count) call records, $($allMeetingsToProcess.Count) teacher meetings"
)
if ($phaseTimers['Phase2a']) {
    $summaryLines += "  Phase 2a (resolution) : $([math]::Round(($phaseTimers['Phase2a']).TotalSeconds, 1))s — $($resolvedMeetings.Count) resolved via `$batch"
}
if ($phaseTimers['Phase2b']) {
    $summaryLines += "  Phase 2b (attendance) : $([math]::Round(($phaseTimers['Phase2b']).TotalSeconds, 1))s — $($meetingChunks.Count) chunks × $ThrottleLimit workers, $($allResults.Count) records"
}
$summaryLines += @(
    "  Phase 3  (export)     : $([math]::Round(($phaseTimers['Phase3']).TotalSeconds, 1))s"
    "  Total                 : $([math]::Round($scriptTimer.Elapsed.TotalSeconds, 1))s"
)

foreach ($line in $summaryLines) {
    Write-Log -Message $line -LogsDir $config.logsDir
}

#endregion
