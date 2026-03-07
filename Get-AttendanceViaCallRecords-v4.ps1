<#
.SYNOPSIS
    Extracts Teams meeting attendance data using Call Records — v4 with fully
    parallelized meeting resolution AND attendance fetching.

.DESCRIPTION
    Builds on v3's chunked-parallel-batch attendance, and extends parallelism to
    Phase 2a (meeting resolution).  In v3, meeting resolution was sequential
    ($batch of 20, one HTTP call at a time).  v4 parallelizes it across workers.

    Architecture:
      Phase 1 : Call Record discovery + OID-first filtering          (same as v3)
      Phase 2a: Meeting resolution via PARALLEL $batch               (NEW — parallel)
                - Split candidates into chunks of 200
                - Each worker sends batches of 20 within its chunk
                - Workers run at ThrottleLimit concurrency
      Phase 2b: Attendance via Chunked-Parallel-Batch                (same as v3)
      Phase 3 : Export to Excel

    Key optimisations over v3:
      - Phase 2a parallelized: ~10x faster meeting resolution
      - Estimated: 35k meetings resolved in ~5-8 min vs ~55 min (v3)

    Estimated HTTP calls at 35k meetings (assuming 1 report per meeting):
      v3 :  ~1,757 sequential (resolution) + ~1,750 parallel (attendance) = ~3,500
      v4 :  ~1,757 parallel   (resolution) + ~1,750 parallel (attendance) = ~3,500
      (Same HTTP count, but Phase 2a is now 10x throughput)

    Requires: PowerShell 7+, CallRecords.Read.All (application).

.PARAMETER TargetDate
    The date to extract attendance for. Default: yesterday in the configured timezone.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER ThrottleLimit
    Maximum concurrent workers for parallel phases. Default: 10.
    Higher values are faster but increase risk of Graph 429 throttling.

.EXAMPLE
    .\Get-AttendanceViaCallRecords-v4.ps1
    .\Get-AttendanceViaCallRecords-v4.ps1 "2026-03-01"
    .\Get-AttendanceViaCallRecords-v4.ps1 -TargetDate "2026-03-01" -ThrottleLimit 15
#>
param(
    [Parameter(Position = 0)]
    [datetime]$TargetDate,
    [string]$ConfigPath = ".\config.json",
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
    $logFile   = Join-Path $LogsDir "callrecords_v4_$(Get-Date -Format 'yyyy-MM-dd').log"

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

    $filePath = Join-Path $OutputDir "callrecords_v4_$($ReportDate.ToString('yyyy-MM-dd')).xlsx"

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
    Get-ChildItem -Path $OutputDir -Filter "callrecords_v4_*.xlsx" -ErrorAction SilentlyContinue |
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

# ── Graph $batch helper with per-item retry (used in main thread) ──
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

Write-Log -Message "=== Call Records approach (v4 — fully parallel) ===" -LogsDir $config.logsDir
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
    param([string]$Uri, [string]$LogsDir)
    $items    = [System.Collections.Generic.List[object]]::new()
    $next     = $Uri
    $pageNum  = 0
    while ($next) {
        $pageNum++
        $resp = Invoke-MgGraphRequest -Uri $next -Method GET -OutputType PSObject
        if ($resp.value) { $items.AddRange($resp.value) }
        $next = $resp.'@odata.nextLink'

        if ($pageNum % 10 -eq 0 -or -not $next) {
            $status = "Phase 1: Fetching call records — page $pageNum, $($items.Count) records so far..."
            Write-Progress -Activity "Teams Attendance Export (v4)" -Status $status -PercentComplete 5
            if ($LogsDir) {
                Write-Log -Message "  Pagination: page $pageNum fetched — $($items.Count) records so far$(if (-not $next) { ' (done)' })" -LogsDir $LogsDir
            }
        }
    }
    return $items
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 1: Discover meetings via Call Records
# ══════════════════════════════════════════════════════════════════════════════

$phase1Timer = [System.Diagnostics.Stopwatch]::StartNew()

Write-Progress -Activity "Teams Attendance Export (v4)" -Status "Phase 1: Discovering meetings via Call Records..." -PercentComplete 5
Write-Log -Message "Phase 1: Listing call records for $($TargetDate.ToString('yyyy-MM-dd'))..." -LogsDir $config.logsDir

$callRecordsUri = "/v1.0/communications/callRecords" +
                  "?`$filter=startDateTime ge $startIso and startDateTime lt $endIso" +
                  "&`$select=id,type,startDateTime,endDateTime,joinWebUrl,organizer"

$allCallRecords = Invoke-MgGraphPaged -Uri $callRecordsUri -LogsDir $config.logsDir

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

Write-Progress -Activity "Teams Attendance Export (v4)" -Status "Phase 1: Oid-first filtering ($($nonTeacherRecords.Count) records)..." -PercentComplete 15

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
# PHASE 2a: Resolve meetings via PARALLEL $batch
#
# v3 sent batches of 20 sequentially (1 HTTP call at a time).
# v4 chunks candidates into groups of 200, then dispatches chunks across
# parallel workers.  Each worker sends its 200 candidates as 10 sequential
# $batch calls of 20, but multiple workers run concurrently.
# At ThrottleLimit=10, this gives ~10x throughput.
# ══════════════════════════════════════════════════════════════════════════════

$resolvedMeetings = [System.Collections.Generic.List[object]]::new()

if ($allMeetingsToProcess.Count -gt 0) {
    $phase2aTimer = [System.Diagnostics.Stopwatch]::StartNew()

    Refresh-TokenIfNeeded
    Write-Progress -Activity "Teams Attendance Export (v4)" -Status "Phase 2a: Resolving meetings via parallel `$batch API..." -PercentComplete 30
    Write-Log -Message "Phase 2a: Resolving $($allMeetingsToProcess.Count) meetings via parallel `$batch API..." -LogsDir $config.logsDir

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
                CallRecordId     = $cr.id
                JoinWebUrl       = $cr.joinWebUrl
                StartDateTime    = [string]$cr.startDateTime
                EndDateTime      = [string]$cr.endDateTime
                CandidateUserIds = @($candidateUserIds)  # Force to array for serialisation
                TeacherName      = $teacherForRow.displayName
                TeacherEmail     = $teacherForRow.mail
                TeacherDept      = $teacherForRow.department
                EncodedJoinUrl   = [Uri]::EscapeDataString($cr.joinWebUrl)
            })
        }
        else {
            Write-Log -Message "  No teacher candidate for call record $($cr.id) — skipping" -Level "WARN" -LogsDir $config.logsDir
        }
    }

    Write-Log -Message "  Meeting candidates prepared: $($meetingCandidates.Count)" -LogsDir $config.logsDir

    # ── Split candidates into chunks for parallel workers ──
    # Each chunk gets 200 candidates → 10 sequential $batch calls of 20 per worker
    $resolutionChunkSize = 200
    $resolutionChunks = [System.Collections.Generic.List[object[]]]::new()
    for ($i = 0; $i -lt $meetingCandidates.Count; $i += $resolutionChunkSize) {
        $end = [Math]::Min($i + $resolutionChunkSize - 1, $meetingCandidates.Count - 1)
        $resolutionChunks.Add(@($meetingCandidates[$i..$end]))
    }

    Write-Log -Message "  Created $($resolutionChunks.Count) resolution chunks of up to $resolutionChunkSize candidates each" -LogsDir $config.logsDir

    # ── Thread-safe progress for Phase 2a ──
    $totalResChunks = $resolutionChunks.Count
    $resProgress = [hashtable]::Synchronized(@{
        Done = 0; Resolved = 0; Failed = 0; Retried = 0; Throttled = 0
    })
    $resLogInterval = [math]::Max(1, [math]::Min(50, [math]::Ceiling($totalResChunks / 10)))

    # ── Parallel resolution: each worker resolves its chunk via $batch ──
    $parallelResults = $resolutionChunks | ForEach-Object -Parallel {
        $chunk = $_

        # Import from parent scope
        $token         = $using:accessToken
        $tokenExpiry   = $using:tokenExpiresAt
        $tenantId      = $using:tokenTenantId
        $clientId      = $using:tokenClientId
        $clientSecret  = $using:tokenClientSecret
        $progress      = $using:resProgress
        $totalChk      = $using:totalResChunks
        $logEvery      = $using:resLogInterval
        $graphBatchUri = "https://graph.microsoft.com/v1.0/`$batch"

        $headers = @{
            Authorization  = "Bearer $token"
            'Content-Type' = 'application/json'
        }

        # ── Local: refresh token if near expiry ──
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
            catch { }
        }

        # ── Local: Send-GraphBatch (same as v3 Phase 2b) ──
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

                for ($i = 0; $i -lt $pending.Count; $i += 20) {
                    $end   = [Math]::Min($i + 19, $pending.Count - 1)
                    $slice = @($pending[$i..$end])

                    $body = @{
                        requests = @($slice | ForEach-Object {
                            @{ id = $_.id; method = $_.method; url = $_.url }
                        })
                    } | ConvertTo-Json -Depth 10

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

        try {
            # ── First pass: try CandidateUserIds[0] for each candidate ──
            $batchRequests = @($chunk | ForEach-Object {
                @{
                    id     = $_.CallRecordId
                    method = "GET"
                    url    = "/users/$($_.CandidateUserIds[0])/onlineMeetings?`$filter=JoinWebUrl eq '$($_.EncodedJoinUrl)'"
                }
            })

            $batchResponses = Send-GraphBatch -Requests $batchRequests -Headers $headers `
                                              -BatchUri $graphBatchUri -ProgressState $progress

            # ── Parse results, collect retries ──
            $retryList = [System.Collections.Generic.List[object]]::new()

            foreach ($mc in $chunk) {
                $resp = $batchResponses[$mc.CallRecordId]
                $meeting = $null

                if ($resp -and $resp.status -eq 200 -and $resp.body.value) {
                    $meetingValues = @($resp.body.value)
                    if ($meetingValues.Count -gt 0) { $meeting = $meetingValues[0] }
                }

                if ($meeting) {
                    $progress.Resolved++
                    # Emit resolved meeting info
                    [PSCustomObject]@{
                        _Type          = 'Resolved'
                        MeetingId      = $meeting.id
                        MeetingSubject = $meeting.subject
                        MeetingUserId  = $mc.CandidateUserIds[0]
                        TeacherName    = $mc.TeacherName
                        TeacherEmail   = $mc.TeacherEmail
                        TeacherDept    = $mc.TeacherDept
                        StartDateTime  = $mc.StartDateTime
                        EndDateTime    = $mc.EndDateTime
                        CallRecordId   = $mc.CallRecordId
                    }
                }
                elseif ($mc.CandidateUserIds.Count -gt 1) {
                    $retryList.Add($mc)
                }
                else {
                    $progress.Failed++
                    [PSCustomObject]@{
                        _Type        = 'Failed'
                        CallRecordId = $mc.CallRecordId
                        Reason       = "No meeting found (status: $(if ($resp) { $resp.status } else { 'null' }))"
                        Candidates   = 1
                    }
                }
            }

            # ── Retry with alternate candidates ──
            if ($retryList.Count -gt 0) {
                $progress.Retried += $retryList.Count

                $retryBatchRequests = @($retryList | ForEach-Object {
                    @{
                        id     = $_.CallRecordId
                        method = "GET"
                        url    = "/users/$($_.CandidateUserIds[1])/onlineMeetings?`$filter=JoinWebUrl eq '$($_.EncodedJoinUrl)'"
                    }
                })

                $retryResponses = Send-GraphBatch -Requests $retryBatchRequests -Headers $headers `
                                                  -BatchUri $graphBatchUri -ProgressState $progress

                foreach ($mc in $retryList) {
                    $resp = $retryResponses[$mc.CallRecordId]
                    $meeting = $null

                    if ($resp -and $resp.status -eq 200 -and $resp.body.value) {
                        $meetingValues = @($resp.body.value)
                        if ($meetingValues.Count -gt 0) { $meeting = $meetingValues[0] }
                    }

                    if ($meeting) {
                        $progress.Resolved++
                        [PSCustomObject]@{
                            _Type          = 'Resolved'
                            MeetingId      = $meeting.id
                            MeetingSubject = $meeting.subject
                            MeetingUserId  = $mc.CandidateUserIds[1]
                            TeacherName    = $mc.TeacherName
                            TeacherEmail   = $mc.TeacherEmail
                            TeacherDept    = $mc.TeacherDept
                            StartDateTime  = $mc.StartDateTime
                            EndDateTime    = $mc.EndDateTime
                            CallRecordId   = $mc.CallRecordId
                        }
                    }
                    else {
                        # Try remaining candidates (3+) sequentially — rare
                        $found = $false
                        for ($ci = 2; $ci -lt $mc.CandidateUserIds.Count; $ci++) {
                            $uid = $mc.CandidateUserIds[$ci]
                            try {
                                $meetingUri = "https://graph.microsoft.com/v1.0/users/$uid/onlineMeetings?`$filter=JoinWebUrl eq '$($mc.EncodedJoinUrl)'"
                                $result = Invoke-RestMethod -Uri $meetingUri -Headers $headers
                                $m = $result.value | Select-Object -First 1
                                if ($m) {
                                    $progress.Resolved++
                                    [PSCustomObject]@{
                                        _Type          = 'Resolved'
                                        MeetingId      = $m.id
                                        MeetingSubject = $m.subject
                                        MeetingUserId  = $uid
                                        TeacherName    = $mc.TeacherName
                                        TeacherEmail   = $mc.TeacherEmail
                                        TeacherDept    = $mc.TeacherDept
                                        StartDateTime  = $mc.StartDateTime
                                        EndDateTime    = $mc.EndDateTime
                                        CallRecordId   = $mc.CallRecordId
                                    }
                                    $found = $true
                                    break
                                }
                            }
                            catch { }
                        }
                        if (-not $found) {
                            $progress.Failed++
                            [PSCustomObject]@{
                                _Type        = 'Failed'
                                CallRecordId = $mc.CallRecordId
                                Reason       = "No meeting found after $($mc.CandidateUserIds.Count) candidates"
                                Candidates   = $mc.CandidateUserIds.Count
                            }
                        }
                    }
                }
            }
        }
        catch {
            # If entire chunk fails, mark all as failed
            $progress.Failed += $chunk.Count
            foreach ($mc in $chunk) {
                [PSCustomObject]@{
                    _Type        = 'Failed'
                    CallRecordId = $mc.CallRecordId
                    Reason       = "Chunk error: $_"
                    Candidates   = $mc.CandidateUserIds.Count
                }
            }
        }
        finally {
            $progress.Done++
            $done = $progress.Done
            $pct  = [math]::Min(100, [int](($done / [math]::Max($totalChk, 1)) * 100))

            Write-Progress -Activity "Phase 2a: Parallel meeting resolution" `
                -Status "$done / $totalChk chunks ($pct%) — resolved: $($progress.Resolved), failed: $($progress.Failed), throttled: $($progress.Throttled)" `
                -PercentComplete $pct

            if ($done % $logEvery -eq 0 -or $done -eq $totalChk) {
                $ts = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                Write-Host "$ts | INFO |   Phase 2a: $done / $totalChk chunks ($pct%) — resolved: $($progress.Resolved), failed: $($progress.Failed), throttled: $($progress.Throttled)"
            }
        }
    } -ThrottleLimit $ThrottleLimit

    Write-Progress -Activity "Phase 2a: Parallel meeting resolution" -Completed

    # ── Collect results from parallel pipeline ──
    $parallelResults = @($parallelResults)

    $failedCount = 0
    foreach ($result in $parallelResults) {
        if ($result._Type -eq 'Resolved') {
            $resolvedMeetings.Add($result)
        }
        elseif ($result._Type -eq 'Failed') {
            $failedCount++
            Write-Log -Message "  No meeting found for call record $($result.CallRecordId): $($result.Reason)" -Level "WARN" -LogsDir $config.logsDir
        }
    }

    $phase2aTimer.Stop()
    $phaseTimers['Phase2a'] = $phase2aTimer.Elapsed
    Write-Log -Message "  Resolved meetings: $($resolvedMeetings.Count) of $($meetingCandidates.Count) in $([math]::Round($phase2aTimer.Elapsed.TotalSeconds, 1))s ($failedCount failed, $($resProgress.Throttled) throttled)" -LogsDir $config.logsDir
}

# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2b: Fetch attendance via Chunked-Parallel-Batch (same as v3)
# ══════════════════════════════════════════════════════════════════════════════

$allResults = @()

if ($resolvedMeetings.Count -gt 0) {
    $phase2bTimer = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Progress -Activity "Teams Attendance Export (v4)" -Status "Phase 2b: Chunked batch attendance for $($resolvedMeetings.Count) meetings..." -PercentComplete 45
    Write-Log -Message "Phase 2b: Fetching attendance for $($resolvedMeetings.Count) meetings via chunked parallel batch (ThrottleLimit $ThrottleLimit)..." -LogsDir $config.logsDir

    # ── Build serializable work items ──
    # In v4, resolved meetings from Phase 2a are already flat scalars
    $skippedNullDates = 0
    $workItems = foreach ($rm in $resolvedMeetings) {
        if ([string]::IsNullOrWhiteSpace($rm.StartDateTime) -or
            [string]::IsNullOrWhiteSpace($rm.EndDateTime)) {
            $skippedNullDates++
            continue
        }
        [PSCustomObject]@{
            MeetingId      = $rm.MeetingId
            MeetingSubject = $rm.MeetingSubject
            MeetingUserId  = $rm.MeetingUserId
            TeacherName    = $rm.TeacherName
            TeacherEmail   = $rm.TeacherEmail
            TeacherDept    = $rm.TeacherDept
            StartDateTime  = [string]$rm.StartDateTime
            EndDateTime    = [string]$rm.EndDateTime
        }
    }
    $workItems = @($workItems)
    if ($skippedNullDates -gt 0) {
        Write-Log -Message "  Skipped $skippedNullDates meeting(s) with null start/end timestamps" -LogsDir $config.logsDir -Level 'WARN'
    }

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
            catch { }
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

                for ($i = 0; $i -lt $pending.Count; $i += 20) {
                    $end   = [Math]::Min($i + 19, $pending.Count - 1)
                    $slice = @($pending[$i..$end])

                    $body = @{
                        requests = @($slice | ForEach-Object {
                            @{ id = $_.id; method = $_.method; url = $_.url }
                        })
                    } | ConvertTo-Json -Depth 10

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

                    $parts      = $key -split '_', 2
                    $meetingIdx = [int]$parts[0]
                    $wi         = $chunk[$meetingIdx]

                    # Guard against null/empty/unparseable timestamps
                    $mStart = [datetime]::MinValue
                    $mEnd   = [datetime]::MinValue
                    if (-not [datetime]::TryParse([string]$wi.StartDateTime, [ref]$mStart) -or
                        -not [datetime]::TryParse([string]$wi.EndDateTime,   [ref]$mEnd)) {
                        continue
                    }

                    $records = $resp.body.attendanceRecords
                    if (-not $records) { continue }

                    foreach ($rec in $records) {
                        $meetingSec = ($mEnd - $mStart).TotalSeconds
                        $attPct = if ($meetingSec -gt 0) {
                                      ($rec.totalAttendanceInSeconds / $meetingSec) * 100
                                  } else { 100 }

                        $joinDt  = $null
                        $leaveDt = $null
                        if ($rec.attendanceIntervals -and @($rec.attendanceIntervals).Count -gt 0) {
                            $tmpJoin  = [datetime]::MinValue
                            $tmpLeave = [datetime]::MinValue
                            if ([datetime]::TryParse([string]$rec.attendanceIntervals[0].joinDateTime, [ref]$tmpJoin)) {
                                $joinDt = $tmpJoin
                            }
                            if ([datetime]::TryParse([string]$rec.attendanceIntervals[-1].leaveDateTime, [ref]$tmpLeave)) {
                                $leaveDt = $tmpLeave
                            }
                        }

                        $lateCutoff = $mStart.AddMinutes($lateMin)
                        $status = if ($attPct -lt $partialPct)                  { "Partial" }
                                  elseif ($joinDt -and $joinDt -gt $lateCutoff) { "Late"    }
                                  else                                          { "Present" }

                        $chunkRecordCount++

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

Write-Progress -Activity "Teams Attendance Export (v4)" -Status "Phase 3: Exporting to Excel..." -PercentComplete 90
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

Write-Progress -Activity "Teams Attendance Export (v4)" -Completed

$scriptTimer.Stop()

# ══════════════════════════════════════════════════════════════════════════════
# Summary
# ══════════════════════════════════════════════════════════════════════════════

$summaryLines = @(
    "=== v4 Performance Summary ==="
    "  Phase 1  (discovery)  : $([math]::Round(($phaseTimers['Phase1']).TotalSeconds, 1))s — $($allCallRecords.Count) call records, $($allMeetingsToProcess.Count) teacher meetings"
)
if ($phaseTimers['Phase2a']) {
    $summaryLines += "  Phase 2a (resolution) : $([math]::Round(($phaseTimers['Phase2a']).TotalSeconds, 1))s — $($resolvedMeetings.Count) resolved via parallel `$batch ($($resolutionChunks.Count) chunks × $ThrottleLimit workers)"
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
