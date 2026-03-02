<#
.SYNOPSIS
    Syncs the teacher list from a Microsoft 365 security group to teachers.json.

.DESCRIPTION
    Authenticates to Microsoft Graph using client credentials and fetches all members
    of the configured teacher security group. Writes the result to teachers.json.

    Uses Invoke-MgGraphRequest for paginated retrieval with automatic token management.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.EXAMPLE
    .\Sync-Teachers.ps1
    .\Sync-Teachers.ps1 -ConfigPath "C:\config\config.json"
#>
param(
    [string]$ConfigPath = ".\config.json"
)

$ErrorActionPreference = "Stop"

# ── Load configuration ──
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

# ── Ensure logs directory exists ──
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
    $logFile   = Join-Path $config.logsDir "sync-teachers_$(Get-Date -Format 'yyyy-MM-dd').log"

    Add-Content -Path $logFile -Value $logLine

    switch ($Level) {
        "ERROR" {
            Write-Warning $logLine
            Add-Content -Path (Join-Path $config.logsDir "errors.log") -Value $logLine
        }
        "WARN"  { Write-Warning $logLine }
        default { Write-Host $logLine }
    }
}

# ── Validate required modules ──
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication")) {
    Write-Error "Microsoft.Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

# ── Connect to Microsoft Graph ──
Write-Log "Connecting to Microsoft Graph..."

$secureSecret = ConvertTo-SecureString $config.clientSecret -AsPlainText -Force
$credential   = New-Object System.Management.Automation.PSCredential($config.clientId, $secureSecret)
Connect-MgGraph -TenantId $config.tenantId -ClientSecretCredential $credential -NoWelcome

Write-Log "Connected to Microsoft Graph (tenant: $($config.tenantId))"

# ── Fetch all group members with pagination ──
# The /microsoft.graph.user cast returns user properties directly, avoiding per-member round-trips.
$uri = "/v1.0/groups/$($config.teacherGroupId)/members/microsoft.graph.user" +
       "?`$select=id,displayName,mail,department,officeLocation&`$top=999"

$teachers = [System.Collections.Generic.List[object]]::new()
$nextUri  = $uri

try {
    while ($nextUri) {
        $resp = Invoke-MgGraphRequest -Uri $nextUri -Method GET -OutputType PSObject

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

    # ── Write output ──
    $outputPath = Join-Path (Split-Path $ConfigPath -Parent) "teachers.json"
    $teachers | ConvertTo-Json -Depth 3 | Set-Content $outputPath -Encoding UTF8

    Write-Log "$($teachers.Count) teachers synced to $outputPath"
}
catch {
    Write-Log "Failed to sync teachers: $_" "ERROR"
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
