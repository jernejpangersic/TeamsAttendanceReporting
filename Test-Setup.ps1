<#
.SYNOPSIS
    Validates that the environment is correctly configured to run the
    Teams Attendance Reporting scripts.

.DESCRIPTION
    Checks PowerShell version, required modules, config.json fields,
    teachers.json, and optionally tests Graph API connectivity.

.PARAMETER ConfigPath
    Path to config.json. Default: .\config.json

.PARAMETER SkipGraphTest
    Skip the Graph API connectivity test.

.EXAMPLE
    .\Test-Setup.ps1
    .\Test-Setup.ps1 -ConfigPath "D:\config.json"
    .\Test-Setup.ps1 -SkipGraphTest
#>
param(
    [string]$ConfigPath = ".\config.json",
    [switch]$SkipGraphTest
)

$pass = 0
$warn = 0
$fail = 0

function Write-Check {
    param([string]$Name, [string]$Status, [string]$Detail)
    switch ($Status) {
        'PASS' {
            $script:pass++
            Write-Host "  [PASS] $Name" -ForegroundColor Green
        }
        'WARN' {
            $script:warn++
            Write-Host "  [WARN] $Name — $Detail" -ForegroundColor Yellow
        }
        'FAIL' {
            $script:fail++
            Write-Host "  [FAIL] $Name — $Detail" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host " Teams Attendance Reporting — Setup Check" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# ──────────────────────────────────────────────────────────────────────────────
# 1. PowerShell Version
# ──────────────────────────────────────────────────────────────────────────────
Write-Host "1. PowerShell Version" -ForegroundColor White
$psVer = $PSVersionTable.PSVersion
if ($psVer.Major -ge 7) {
    Write-Check "PowerShell $psVer" "PASS"
} elseif ($psVer.Major -ge 5) {
    Write-Check "PowerShell $psVer" "WARN" "v5.x works for basic scripts but v7+ is required for parallel scripts (v3-v5). Install with: winget install Microsoft.PowerShell"
} else {
    Write-Check "PowerShell $psVer" "FAIL" "PowerShell 7+ is required. Install with: winget install Microsoft.PowerShell"
}

# ──────────────────────────────────────────────────────────────────────────────
# 2. Execution Policy
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "2. Execution Policy" -ForegroundColor White
$epolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($epolicy -in @('RemoteSigned', 'Unrestricted', 'Bypass')) {
    Write-Check "Execution policy: $epolicy" "PASS"
} elseif ((Get-ExecutionPolicy -Scope Process) -in @('RemoteSigned', 'Unrestricted', 'Bypass')) {
    Write-Check "Execution policy: $(Get-ExecutionPolicy -Scope Process) (Process scope)" "PASS"
} else {
    Write-Check "Execution policy: $epolicy" "WARN" "Scripts may be blocked. Run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
}

# ──────────────────────────────────────────────────────────────────────────────
# 3. Required Modules
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "3. PowerShell Modules" -ForegroundColor White

$modules = @(
    @{ Name = 'Microsoft.Graph.Authentication'; Install = 'Install-Module Microsoft.Graph -Scope CurrentUser'; Required = $true }
    @{ Name = 'ImportExcel';                    Install = 'Install-Module ImportExcel -Scope CurrentUser';     Required = $true }
    @{ Name = 'MicrosoftTeams';                 Install = 'Install-Module MicrosoftTeams -Scope CurrentUser';  Required = $false }
)

$missingModules = @()
foreach ($mod in $modules) {
    $installed = Get-Module -ListAvailable -Name $mod.Name | Select-Object -First 1
    if ($installed) {
        Write-Check "$($mod.Name) v$($installed.Version)" "PASS"
    } else {
        $missingModules += $mod
        if ($mod.Required) {
            Write-Check "$($mod.Name)" "FAIL" "Not installed"
        } else {
            Write-Check "$($mod.Name)" "WARN" "Not installed (only needed for one-time policy setup)"
        }
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host ""
    Write-Host "  Missing modules:" -ForegroundColor Yellow
    foreach ($mod in $missingModules) {
        $tag = if ($mod.Required) { "required" } else { "optional" }
        Write-Host "    - $($mod.Name) ($tag)" -ForegroundColor Yellow
    }
    $answer = Read-Host "  Install missing modules now? (Y/N)"
    if ($answer -match '^[Yy]') {
        foreach ($mod in $missingModules) {
            Write-Host "  Installing $($mod.Name)..." -ForegroundColor Cyan
            try {
                Invoke-Expression $mod.Install
                Write-Check "$($mod.Name) installed" "PASS"
                # Adjust counters: undo the previous FAIL/WARN and count as PASS
                $script:pass++
                if ($mod.Required) { $script:fail-- } else { $script:warn-- }
            }
            catch {
                Write-Check "$($mod.Name) install" "FAIL" "$_"
            }
        }
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# 4. config.json
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "4. Configuration (config.json)" -ForegroundColor White

$configOk = $false
if (Test-Path $ConfigPath) {
    Write-Check "config.json exists at $ConfigPath" "PASS"
    try {
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $configOk = $true

        $requiredFields = @('tenantId', 'clientId', 'clientSecret', 'teacherGroupId')
        foreach ($field in $requiredFields) {
            $val = $config.$field
            if ([string]::IsNullOrWhiteSpace($val) -or $val -match '^<.*>$') {
                Write-Check "config.$field" "FAIL" "Missing or still set to placeholder value"
                $configOk = $false
            } else {
                # Mask secrets in output
                $display = if ($field -eq 'clientSecret') { "$($val.Substring(0, [math]::Min(4, $val.Length)))****" } else { $val }
                Write-Check "config.$field = $display" "PASS"
            }
        }

        $optionalFields = @('outputDir', 'logsDir', 'timezone')
        foreach ($field in $optionalFields) {
            $val = $config.$field
            if ([string]::IsNullOrWhiteSpace($val)) {
                Write-Check "config.$field" "WARN" "Not set (will use default)"
            } else {
                Write-Check "config.$field = $val" "PASS"
            }
        }
    }
    catch {
        Write-Check "config.json parse" "FAIL" "Invalid JSON: $_"
    }
} else {
    Write-Check "config.json" "FAIL" "Not found at $ConfigPath. Run: Copy-Item config.sample.json config.json"
}

# ──────────────────────────────────────────────────────────────────────────────
# 5. teachers.json
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "5. Teacher List (teachers.json)" -ForegroundColor White

$teacherPath = Join-Path (Split-Path $ConfigPath -Parent) "teachers.json"
if (Test-Path $teacherPath) {
    try {
        $teachers = Get-Content $teacherPath -Raw | ConvertFrom-Json
        $count = @($teachers).Count
        if ($count -gt 0) {
            Write-Check "teachers.json: $count teacher(s)" "PASS"
        } else {
            Write-Check "teachers.json" "WARN" "File exists but contains 0 teachers"
        }
    }
    catch {
        Write-Check "teachers.json parse" "FAIL" "Invalid JSON: $_"
    }
} else {
    Write-Check "teachers.json" "FAIL" "Not found. Run: .\Sync-Teachers.ps1"
}

# ──────────────────────────────────────────────────────────────────────────────
# 6. Output / Logs Directories
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "6. Directories" -ForegroundColor White

$outDir  = if ($config.outputDir) { $config.outputDir } else { "./output" }
$logsDir = if ($config.logsDir)   { $config.logsDir }   else { "./logs" }

foreach ($dir in @($outDir, $logsDir)) {
    if (Test-Path $dir) {
        Write-Check "$dir exists" "PASS"
    } else {
        Write-Check "$dir" "WARN" "Does not exist yet (will be created automatically on first run)"
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# 7. Graph API Connectivity (optional)
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "7. Graph API Connectivity" -ForegroundColor White

if ($SkipGraphTest) {
    Write-Check "Skipped (use without -SkipGraphTest to test)" "WARN" "Graph connectivity not verified"
} elseif (-not $configOk) {
    Write-Check "Skipped" "WARN" "config.json has errors — fix those first"
} else {
    try {
        $tokenBody = @{
            client_id     = $config.clientId
            client_secret = $config.clientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }
        $tokenResp = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$($config.tenantId)/oauth2/v2.0/token" `
            -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded"

        if ($tokenResp.access_token) {
            Write-Check "OAuth token acquired" "PASS"

            # Quick test: read the teacher group
            $headers = @{ Authorization = "Bearer $($tokenResp.access_token)" }
            try {
                $groupResp = Invoke-RestMethod `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$($config.teacherGroupId)?`$select=id,displayName,membershipRule" `
                    -Headers $headers
                Write-Check "Teacher group: $($groupResp.displayName) ($($config.teacherGroupId))" "PASS"
            }
            catch {
                $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
                if ($sc -eq 403) {
                    Write-Check "Teacher group lookup" "FAIL" "403 Forbidden — app may be missing Group.Read.All permission or admin consent"
                } elseif ($sc -eq 404) {
                    Write-Check "Teacher group lookup" "FAIL" "404 Not Found — teacherGroupId '$($config.teacherGroupId)' does not exist"
                } else {
                    Write-Check "Teacher group lookup" "FAIL" "HTTP $sc — $_"
                }
            }
        }
    }
    catch {
        $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        if ($sc -eq 401) {
            Write-Check "OAuth token" "FAIL" "401 Unauthorized — check clientId and clientSecret in config.json"
        } elseif ($sc -eq 400) {
            Write-Check "OAuth token" "FAIL" "400 Bad Request — check tenantId in config.json"
        } else {
            Write-Check "OAuth token" "FAIL" "$_"
        }
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Summary
# ──────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host " Results: $pass passed, $warn warnings, $fail failed" -ForegroundColor $(if ($fail -gt 0) { 'Red' } elseif ($warn -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

if ($fail -gt 0) {
    Write-Host "Fix the FAIL items above before running the attendance scripts." -ForegroundColor Red
} elseif ($warn -gt 0) {
    Write-Host "Environment looks mostly ready. Review the WARN items above." -ForegroundColor Yellow
} else {
    Write-Host "Environment is fully configured. You're ready to go!" -ForegroundColor Green
    Write-Host "  Next step: .\Get-AttendanceViaCallRecords-v5.ps1" -ForegroundColor Gray
}
