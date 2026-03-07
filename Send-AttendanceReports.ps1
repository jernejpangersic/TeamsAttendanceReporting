<#
.SYNOPSIS
    Sends per-department attendance Excel files to school IT admins via Microsoft Graph.

.DESCRIPTION
    Reads department Excel files from a report folder (produced by Split-AttendanceByDepartment.ps1),
    matches each file to recipients defined in recipients.json, and sends the file as an email
    attachment using the Microsoft Graph sendMail API.

    The email body is loaded from an HTML template (email-template.html) with placeholder
    substitution for Department, Date, and RowCount.

.PARAMETER ReportDir
    Path to the folder containing per-department Excel files. If omitted, auto-detects
    the most recently created date folder under .\output.

.PARAMETER RecipientsFile
    Path to the recipients JSON file. Default: .\recipients.json

.PARAMETER TemplatePath
    Path to the HTML email body template. Default: .\email-template.html

.PARAMETER Subject
    Email subject line. Use {{Date}} as a placeholder for the report date.
    Default: "Attendance Report - {{Date}}"

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER WhatIf
    Preview what would be sent without actually sending.

.EXAMPLE
    .\Send-AttendanceReports.ps1 -ReportDir .\output\callrecords_v5_2026-03-02

.EXAMPLE
    .\Send-AttendanceReports.ps1 -WhatIf

.EXAMPLE
    .\Send-AttendanceReports.ps1 -Subject "Weekly Report - {{Date}}" -TemplatePath .\custom-email.html
#>

param(
    [string]$ReportDir,
    [string]$RecipientsFile = ".\recipients.json",
    [string]$TemplatePath   = ".\email-template.html",
    [string]$Subject        = "Attendance Report - {{Date}}",
    [string]$ConfigPath     = ".\config.json",
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Load configuration ──
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    return
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

if (-not $config.senderEmail) {
    Write-Error "config.json is missing the 'senderEmail' field. Add it and try again."
    return
}

# ── Ensure logs directory ──
if (-not (Test-Path $config.logsDir)) {
    New-Item -ItemType Directory -Path $config.logsDir -Force | Out-Null
}

# ── Logging helper ──
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $logLine   = "$timestamp | $Level | $Message"
    $logFile   = Join-Path $config.logsDir "send-reports_$(Get-Date -Format 'yyyy-MM-dd').log"

    Add-Content -Path $logFile -Value $logLine

    switch ($Level) {
        "ERROR" { Write-Warning $logLine }
        "WARN"  { Write-Warning $logLine }
        default { Write-Host $logLine }
    }
}

# ── Load recipients ──
if (-not (Test-Path $RecipientsFile)) {
    Write-Error "Recipients file not found: $RecipientsFile"
    return
}
$recipientsRaw = Get-Content $RecipientsFile -Encoding UTF8 | ConvertFrom-Json
$recipients = @{}
foreach ($prop in $recipientsRaw.PSObject.Properties) {
    if ($prop.Name -like '_*') { continue }   # skip _example, _instructions, etc.
    $emails = @(@($prop.Value) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($emails.Count -gt 0) {
        $recipients[$prop.Name] = $emails
    }
}

if ($recipients.Count -eq 0) {
    Write-Error "No valid recipients found in $RecipientsFile. Fill in email addresses and try again."
    return
}

# ── Load email template ──
if (-not (Test-Path $TemplatePath)) {
    Write-Error "Email template not found: $TemplatePath"
    return
}
$templateHtml = Get-Content $TemplatePath -Encoding UTF8 -Raw

# ── Determine report directory ──
if (-not $ReportDir) {
    # Auto-detect: find the most recently created subfolder in .\output
    $outputBase = $config.outputDir
    if (-not (Test-Path $outputBase)) {
        Write-Error "Output directory not found: $outputBase"
        return
    }
    $latestFolder = Get-ChildItem -Path $outputBase -Directory |
        Sort-Object CreationTime -Descending |
        Select-Object -First 1

    if (-not $latestFolder) {
        Write-Error "No subfolders found in $outputBase. Run Split-AttendanceByDepartment.ps1 first."
        return
    }
    $ReportDir = $latestFolder.FullName
    Write-Log "Auto-detected report folder: $ReportDir"
}

if (-not (Test-Path $ReportDir)) {
    Write-Error "Report directory not found: $ReportDir"
    return
}

# ── Extract date from folder name ──
$folderName = Split-Path $ReportDir -Leaf
if ($folderName -match '(\d{4}-\d{2}-\d{2})') {
    $reportDate = $Matches[1]
} else {
    $reportDate = $folderName
}

# ── Discover Excel files ──
$excelFiles = @(Get-ChildItem -Path $ReportDir -Filter *.xlsx -File)
if ($excelFiles.Count -eq 0) {
    Write-Error "No Excel files found in $ReportDir"
    return
}

Write-Log "Found $($excelFiles.Count) Excel file(s) in $ReportDir"
Write-Log "Recipients configured for $($recipients.Count) department(s)"
Write-Host ""

# ── Connect to Microsoft Graph ──
if (-not $WhatIf) {
    if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication")) {
        Write-Error "Microsoft.Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
        return
    }

    Write-Log "Connecting to Microsoft Graph..."
    $secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
    $credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
    Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome
    Write-Log "Connected to Microsoft Graph"
}

# ── Send emails ──
$sent    = 0
$skipped = 0
$failed  = 0

foreach ($file in $excelFiles) {
    $deptName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    if (-not $recipients.ContainsKey($deptName)) {
        Write-Log "SKIP: No recipients for '$deptName'" "WARN"
        $skipped++
        continue
    }

    $toAddresses = $recipients[$deptName]

    # Get row count from the Excel file
    $rowCount = (Import-Excel -Path $file.FullName | Measure-Object).Count

    # Build email body from template
    $body = $templateHtml `
        -replace '{{Department}}', $deptName `
        -replace '{{Date}}', $reportDate `
        -replace '{{RowCount}}', $rowCount

    # Build subject
    $emailSubject = $Subject -replace '{{Date}}', $reportDate

    # Read file as base64 for attachment
    $fileBytes  = [System.IO.File]::ReadAllBytes($file.FullName)
    $base64File = [System.Convert]::ToBase64String($fileBytes)

    # Build Graph sendMail payload
    $toRecipients = @($toAddresses | ForEach-Object {
        @{ emailAddress = @{ address = $_ } }
    })

    $mailPayload = @{
        message = @{
            subject      = $emailSubject
            body         = @{
                contentType = "HTML"
                content     = $body
            }
            toRecipients = $toRecipients
            attachments  = @(
                @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    name           = $file.Name
                    contentType    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    contentBytes   = $base64File
                }
            )
        }
        saveToSentItems = $false
    }

    if ($WhatIf) {
        Write-Host "[WhatIf] Would send '$deptName' ($rowCount rows) to: $($toAddresses -join ', ')" -ForegroundColor Cyan
        $sent++
        continue
    }

    try {
        $uri = "/v1.0/users/$($config.senderEmail)/sendMail"
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body ($mailPayload | ConvertTo-Json -Depth 10) -ContentType "application/json"
        Write-Log "SENT: '$deptName' ($rowCount rows) -> $($toAddresses -join ', ')"
        $sent++
    }
    catch {
        Write-Log "FAIL: '$deptName' -> $($toAddresses -join ', '): $_" "ERROR"
        $failed++
    }
}

# ── Summary ──
Write-Host ""
Write-Log "Complete. Sent: $sent | Skipped (no recipients): $skipped | Failed: $failed"

if (-not $WhatIf) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
