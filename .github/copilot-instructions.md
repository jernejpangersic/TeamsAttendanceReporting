# Copilot Instructions for teamsReporting

## README maintenance

When modifying any `.ps1` script, review `README.md` and update it if any of the following changed:
- Script parameters (added, removed, renamed, or default values changed)
- Prerequisites or required modules
- Setup steps or configuration fields
- Usage examples
- Project structure (new or removed files)
- Output format or columns

## Test-Setup.ps1

When adding new prerequisites or configuration fields, update `Test-Setup.ps1` to validate them.

## Code conventions

- All attendance scripts target **PowerShell 7+** (for `ForEach-Object -Parallel`).
- Use `Write-Log` for structured logging (not bare `Write-Host` for operational messages).
- Graph API calls must include retry logic for 429/503/504 status codes.
- DateTime values from external sources must use `[datetime]::TryParse()`, never direct `[datetime]` casts.

## Testing

After modifying any attendance script (`Get-Attendance*.ps1`, `Get-AttendanceViaCallRecords*.ps1`), run a PowerShell syntax parse check to catch errors before the user does:
```powershell
$errors = $null
[System.Management.Automation.Language.Parser]::ParseFile('<script-path>', [ref]$null, [ref]$errors)
```
Report any parse errors and fix them before finishing.
