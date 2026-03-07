<#
.SYNOPSIS
    Splits attendance Excel file(s) into separate files per department.

.DESCRIPTION
    Reads one or more attendance Excel files and creates one Excel file per unique
    Department value. Files are saved in subfolders under .\output.

    When a single file is provided (or only one file exists for a date), the subfolder
    is named by date (e.g. 2026-03-02\). When multiple files share the same date,
    each gets its own subfolder using the full filename stem (e.g. callrecords_v4_2026-03-02\).

.PARAMETER ExcelPath
    Optional. Path to a specific source Excel file. If omitted, all *.xlsx files
    in .\output that contain a YYYY-MM-DD date pattern are processed.

.EXAMPLE
    .\Split-AttendanceByDepartment.ps1 -ExcelPath .\output\callrecords_v5_2026-03-02.xlsx

.EXAMPLE
    .\Split-AttendanceByDepartment.ps1
    # Processes all dated Excel files in .\output
#>

param(
    [string]$ExcelPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Ensure ImportExcel module ──
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "The ImportExcel module is required. Install it with: Install-Module ImportExcel -Scope CurrentUser"
    return
}
Import-Module ImportExcel

# ── Characters illegal in Windows filenames ──
$illegalChars = '[\\/:*?"<>|]'

# ── Conditional formatting (same as the main script) ──
$conditionalText = @(
    New-ConditionalText -Text "Present" -BackgroundColor LightGreen
    New-ConditionalText -Text "Late"    -BackgroundColor Yellow
    New-ConditionalText -Text "Partial" -BackgroundColor Orange
    New-ConditionalText -Text "Absent"  -BackgroundColor Red
)

# ── Determine which files to process ──
if ($ExcelPath) {
    if (-not (Test-Path $ExcelPath)) {
        Write-Error "File not found: $ExcelPath"
        return
    }
    $filesToProcess = @(Get-Item $ExcelPath)
    $baseDir = Split-Path $ExcelPath -Parent
} else {
    $baseDir = ".\output"
    if (-not (Test-Path $baseDir)) {
        Write-Error "Output directory not found: $baseDir"
        return
    }
    $filesToProcess = @(Get-ChildItem -Path $baseDir -Filter *.xlsx -File |
        Where-Object { $_.BaseName -match '\d{4}-\d{2}-\d{2}' })
    if ($filesToProcess.Count -eq 0) {
        Write-Warning "No Excel files with a date pattern found in $baseDir"
        return
    }
    Write-Host "Found $($filesToProcess.Count) Excel file(s) in $baseDir"
    Write-Host ""
}

# ── Group files by their date to decide folder naming ──
$filesByDate = $filesToProcess | Group-Object {
    if ($_.BaseName -match '(\d{4}-\d{2}-\d{2})') { $Matches[1] } else { 'unknown' }
}

# ── Build a lookup: for each file, what folder name should we use? ──
$folderMap = @{}
foreach ($dateGroup in $filesByDate) {
    if ($dateGroup.Count -eq 1) {
        # Only one file for this date → folder = just the date
        $folderMap[$dateGroup.Group[0].FullName] = $dateGroup.Name
    } else {
        # Multiple files for the same date → folder = full filename stem
        foreach ($f in $dateGroup.Group) {
            $folderMap[$f.FullName] = $f.BaseName
        }
    }
}

# ── Process each file ──
foreach ($file in $filesToProcess) {
    $folderName = $folderMap[$file.FullName]
    $outputDir  = Join-Path $baseDir $folderName

    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    Write-Host "Reading $($file.Name) ..."
    $rows = Import-Excel -Path $file.FullName

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Warning "  No data found — skipping."
        continue
    }

    Write-Host "  $($rows.Count) rows imported."

    $groups = $rows | Group-Object -Property Department
    Write-Host "  $($groups.Count) department(s) found."

    foreach ($group in $groups) {
        $deptName = if ([string]::IsNullOrWhiteSpace($group.Name)) { '_No Department' } else { $group.Name }
        $safeName = ($deptName -replace $illegalChars, '_').Trim()
        $outPath  = Join-Path $outputDir "$safeName.xlsx"

        $group.Group | Export-Excel -Path $outPath -WorksheetName "Attendance" `
            -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow `
            -ConditionalText $conditionalText

        Write-Host "    [$($group.Group.Count) rows] -> $outPath"
    }

    Write-Host "  Done. $($groups.Count) file(s) written to $outputDir"
    Write-Host ""
}

Write-Host "All files processed."
