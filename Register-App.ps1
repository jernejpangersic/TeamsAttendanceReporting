<#
.SYNOPSIS
    Registers the TeamsAttendanceReporting Azure AD application with required Graph permissions.

.DESCRIPTION
    Creates an app registration in Entra ID, assigns the necessary Microsoft Graph application
    permissions, generates a client secret, and outputs the values needed for config.json.

    Run this interactively as a tenant admin. After running, grant admin consent in the Entra ID
    portal and create the application access policy via Teams PowerShell.

.PARAMETER AppDisplayName
    Display name for the app registration. Default: "TeamsAttendanceReporting"

.EXAMPLE
    .\Register-App.ps1
    .\Register-App.ps1 -AppDisplayName "MyCustomAppName"
#>
param(
    [string]$AppDisplayName = "TeamsAttendanceReporting"
)

$ErrorActionPreference = "Stop"

# ── Verify module ──
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Applications")) {
    Write-Error "Microsoft.Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Write-Host "=== Azure AD App Registration for Teams Attendance Reporting ===" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Connect with admin permissions ──
Write-Host "Connecting to Microsoft Graph (requires Application.ReadWrite.All)..."
Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome

$context = Get-MgContext
if (-not $context) {
    Write-Error "Failed to connect to Microsoft Graph."
    exit 1
}
Write-Host "Connected to tenant: $($context.TenantId)" -ForegroundColor Green

# ── Step 2: Look up Microsoft Graph service principal for permission IDs ──
Write-Host "Looking up Microsoft Graph permission IDs..."
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

if (-not $graphSp) {
    Write-Error "Could not find Microsoft Graph service principal. Ensure Microsoft Graph is enabled in the tenant."
    exit 1
}

$permissionNames = @(
    "OnlineMeetingArtifact.Read.All",
    "OnlineMeetings.Read.All",
    "User.Read.All",
    "Group.Read.All",
    "Calendars.Read",
    "CallRecords.Read.All"
)

$resourceAccess = @()
foreach ($name in $permissionNames) {
    $perm = $graphSp.AppRoles | Where-Object { $_.Value -eq $name }
    if ($perm) {
        $resourceAccess += @{
            Id   = $perm.Id.ToString()
            Type = "Role"   # "Role" = Application permission
        }
        Write-Host "  Found: $name ($($perm.Id))" -ForegroundColor Gray
    }
    else {
        Write-Warning "  Permission '$name' not found — verify it is available in your tenant."
    }
}

if ($resourceAccess.Count -ne $permissionNames.Count) {
    Write-Warning "Not all permissions were found. The app may not work correctly."
}

# ── Step 3: Create the application ──
Write-Host "Creating application '$AppDisplayName'..."
$appBody = @{
    DisplayName            = $AppDisplayName
    SignInAudience         = "AzureADMyOrg"
    RequiredResourceAccess = @(
        @{
            ResourceAppId  = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
            ResourceAccess = $resourceAccess
        }
    )
}

$app = New-MgApplication -BodyParameter $appBody
Write-Host "Application created — AppId: $($app.AppId)" -ForegroundColor Green

# ── Step 4: Create a client secret (1-year validity) ──
Write-Host "Creating client secret (1-year validity)..."
$secretParams = @{
    PasswordCredential = @{
        DisplayName = "$AppDisplayName-Secret"
        EndDateTime = (Get-Date).AddYears(1)
    }
}
$secret = Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams

# ── Step 5: Output results ──
Write-Host ""
Write-Host "=== Registration Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "App ID (clientId):  $($app.AppId)"
Write-Host "Tenant ID:          $($context.TenantId)"
Write-Host "Client Secret:      $($secret.SecretText)"
Write-Host "Secret Expiry:      $($secret.EndDateTime)"
Write-Host ""
Write-Host "=== NEXT STEPS ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Update config.json with these values:" -ForegroundColor White
Write-Host "   tenantId:     $($context.TenantId)"
Write-Host "   clientId:     $($app.AppId)"
Write-Host "   clientSecret: $($secret.SecretText)"
Write-Host ""
Write-Host "2. Grant admin consent in Entra ID portal:" -ForegroundColor White
Write-Host "   https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)"
Write-Host ""
Write-Host "3. Create and assign application access policy (Teams PowerShell):" -ForegroundColor White
Write-Host ""
Write-Host "   New-CsApplicationAccessPolicy -Identity `"AttendanceReportingPolicy`" ``" -ForegroundColor Cyan
Write-Host "       -AppIds `"$($app.AppId)`" ``" -ForegroundColor Cyan
Write-Host "       -Description `"Allow $AppDisplayName to read meetings and attendance`"" -ForegroundColor Cyan
Write-Host ""
Write-Host "   Grant-CsApplicationAccessPolicy -PolicyName `"AttendanceReportingPolicy`" -Global" -ForegroundColor Cyan
Write-Host ""
Write-Host "4. Wait 30 minutes for policy propagation before running scripts." -ForegroundColor White
Write-Host ""
Write-Host "IMPORTANT: Save the client secret now — it cannot be retrieved later!" -ForegroundColor Red

Disconnect-MgGraph
