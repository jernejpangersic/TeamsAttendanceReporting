<#
.SYNOPSIS
    Creates test Teams meetings via the Microsoft Graph OnlineMeetings API.

.DESCRIPTION
    Generates one or more online meetings for a chosen teacher (organizer) so you can
    test the Get-Attendance.ps1 attendance-extraction pipeline end-to-end.

    Workflow:
      1. Script creates meetings via POST /users/{userId}/onlineMeetings.
      2. Share the join URLs with a few people (or open them yourself in multiple
         browser profiles) so they join briefly.
      3. After the meetings end, run Get-Attendance.ps1 to pull attendance reports.

    PREREQUISITES:
      - The app registration must have **OnlineMeeting.ReadWrite.All** (Application)
        permission. Grant admin consent in the Entra ID portal.
      - The application access policy must cover the organizer user.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER OrganizerEmail
    E-mail (UPN) of the teacher who will be the meeting organizer.
    If omitted, the script lists teachers from teachers.json and lets you pick one.

.PARAMETER MeetingCount
    Number of test meetings to create. Default: 3

.PARAMETER DurationMinutes
    Duration of each meeting in minutes. Default: 30

.PARAMETER StartOffset
    Minutes from *now* for the first meeting to start. Default: 5 (starts in 5 minutes).
    Use 0 or negative values to create meetings in the past (useful for testing, but
    note that attendance reports only exist if people actually joined).

.PARAMETER GapMinutes
    Gap between consecutive meetings in minutes. Default: 5

.PARAMETER SubjectPrefix
    Prefix for the meeting subject. Default: "Test Meeting"

.PARAMETER ParticipantEmails
    Optional array of e-mail addresses to invite as participants.
    These are set in the meeting's Participants.Attendees list.

.EXAMPLE
    # Interactive — pick an organizer, create 3 meetings starting in 5 minutes
    .\New-TestMeetings.ps1

.EXAMPLE
    # Fully automated — specific organizer, 2 meetings, with participants
    .\New-TestMeetings.ps1 -OrganizerEmail "AllanD@M365EDU573063.OnMicrosoft.com" `
                           -MeetingCount 2 -DurationMinutes 15 `
                           -ParticipantEmails @("CaraC@M365EDU573063.OnMicrosoft.com","GeorgeP@M365EDU573063.OnMicrosoft.com")
#>
param(
    [string]$ConfigPath        = ".\config.json",
    [string]$OrganizerEmail,
    [int]$MeetingCount         = 3,
    [int]$DurationMinutes      = 30,
    [int]$StartOffset          = 5,
    [int]$GapMinutes           = 5,
    [string]$SubjectPrefix     = "Test Meeting",
    [string[]]$ParticipantEmails
)

$ErrorActionPreference = "Stop"

#region ── Load Config & Connect ──

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

# Load teachers
$teachersPath = Join-Path (Split-Path $ConfigPath -Parent) "teachers.json"
if (-not (Test-Path $teachersPath)) {
    Write-Error "teachers.json not found at $teachersPath. Run Sync-Teachers.ps1 first."
    exit 1
}
$teachers = Get-Content $teachersPath | ConvertFrom-Json

# ── Pick organizer ──
if (-not $OrganizerEmail) {
    Write-Host "`nAvailable teachers:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $teachers.Count; $i++) {
        Write-Host "  [$i] $($teachers[$i].displayName) ($($teachers[$i].mail))"
    }
    $choice = Read-Host "`nEnter the number of the organizer"
    $organizer = $teachers[[int]$choice]
}
else {
    $organizer = $teachers | Where-Object { $_.mail -ieq $OrganizerEmail }
    if (-not $organizer) {
        Write-Error "Organizer '$OrganizerEmail' not found in teachers.json"
        exit 1
    }
}

Write-Host "`nOrganizer: $($organizer.displayName) ($($organizer.mail))" -ForegroundColor Green

# ── Connect to Graph ──
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
Write-Host "Connected to Microsoft Graph" -ForegroundColor Green

#endregion

#region ── Build & Create Meetings ──

# Resolve timezone
try {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($config.timezone)
}
catch {
    Write-Warning "Could not resolve timezone '$($config.timezone)', falling back to UTC"
    $tz = [System.TimeZoneInfo]::Utc
}

$nowUtc       = [datetime]::UtcNow
$firstStartUtc = $nowUtc.AddMinutes($StartOffset)

$createdMeetings = @()

for ($i = 0; $i -lt $MeetingCount; $i++) {
    $meetingStartUtc = $firstStartUtc.AddMinutes($i * ($DurationMinutes + $GapMinutes))
    $meetingEndUtc   = $meetingStartUtc.AddMinutes($DurationMinutes)

    $subject = "$SubjectPrefix $($i + 1)"

    # Build the request body
    $body = @{
        subject        = $subject
        startDateTime  = $meetingStartUtc.ToString("yyyy-MM-ddTHH:mm:ss.0000000Z")
        endDateTime    = $meetingEndUtc.ToString("yyyy-MM-ddTHH:mm:ss.0000000Z")
        isEntryExitAnnounced = $true
        allowedPresenters    = "everyone"
        lobbyBypassSettings  = @{
            scope       = "everyone"
            isDialInBypassEnabled = $true
        }
    }

    # Add participants if provided
    if ($ParticipantEmails -and $ParticipantEmails.Count -gt 0) {
        $attendees = @()
        foreach ($email in $ParticipantEmails) {
            $attendees += @{
                upn  = $email
                role = "attendee"
            }
        }
        $body.participants = @{
            attendees = $attendees
        }
    }

    $bodyJson = $body | ConvertTo-Json -Depth 5

    Write-Host "`nCreating: '$subject'" -ForegroundColor Yellow
    Write-Host "  Start : $($meetingStartUtc.ToString('yyyy-MM-dd HH:mm')) UTC"
    Write-Host "  End   : $($meetingEndUtc.ToString('yyyy-MM-dd HH:mm')) UTC"

    try {
        $result = Invoke-MgGraphRequest `
            -Uri "/v1.0/users/$($organizer.id)/onlineMeetings" `
            -Method POST `
            -Body $bodyJson `
            -ContentType "application/json" `
            -OutputType PSObject

        $localStart = [System.TimeZoneInfo]::ConvertTimeFromUtc($meetingStartUtc, $tz)
        $localEnd   = [System.TimeZoneInfo]::ConvertTimeFromUtc($meetingEndUtc, $tz)

        $info = [PSCustomObject]@{
            '#'         = $i + 1
            Subject     = $subject
            MeetingId   = $result.id
            JoinUrl     = $result.joinWebUrl
            StartUTC    = $meetingStartUtc.ToString('yyyy-MM-dd HH:mm')
            EndUTC      = $meetingEndUtc.ToString('yyyy-MM-dd HH:mm')
            StartLocal  = $localStart.ToString('yyyy-MM-dd HH:mm')
            EndLocal    = $localEnd.ToString('yyyy-MM-dd HH:mm')
            Organizer   = $organizer.mail
        }

        $createdMeetings += $info

        Write-Host "  OK — Meeting ID: $($result.id)" -ForegroundColor Green
        Write-Host "  Join URL: $($result.joinWebUrl)" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "  FAILED to create '$subject': $($_.Exception.Message)"
        if ($_.Exception.Message -match "OnlineMeeting\.ReadWrite\.All|Authorization_RequestDenied|Forbidden") {
            Write-Host "`n  ** The app registration likely needs the OnlineMeeting.ReadWrite.All permission. **" -ForegroundColor Red
            Write-Host "  Add it in Entra ID > App registrations > API permissions, then grant admin consent.`n" -ForegroundColor Red
        }
    }
}

#endregion

#region ── Summary ──

if ($createdMeetings.Count -gt 0) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  $($createdMeetings.Count) meeting(s) created successfully" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $createdMeetings | Format-Table -Property '#', Subject, StartLocal, EndLocal, JoinUrl -AutoSize -Wrap

    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Share the Join URLs above with a few people (or open in multiple browser profiles)."
    Write-Host "  2. Have them join the meeting(s) for at least 1-2 minutes."
    Write-Host "  3. After the meeting ends, wait ~5 minutes for Teams to finalize attendance data."
    Write-Host "  4. Run Get-Attendance.ps1 to extract the attendance reports:"
    Write-Host ""

    $targetDate = $createdMeetings[0].StartUTC.Substring(0, 10)
    Write-Host "     .\Get-Attendance.ps1 -TargetDate `"$targetDate`"" -ForegroundColor White
    Write-Host ""

    # Optionally export to a small JSON for reference
    $jsonPath = Join-Path $config.outputDir "test_meetings_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $createdMeetings | ConvertTo-Json -Depth 3 | Set-Content -Path $jsonPath -Encoding UTF8
    Write-Host "  Meeting details saved to: $jsonPath" -ForegroundColor Gray
}
else {
    Write-Host "`nNo meetings were created." -ForegroundColor Red
}

#endregion
