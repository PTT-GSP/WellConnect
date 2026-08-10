# Set UTF8 output encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Find the Excel file
$xlsxFiles = Get-ChildItem -Path . -Filter "*.xlsx"
if ($xlsxFiles.Count -eq 0) {
    Write-Error "Error: No Excel (.xlsx) file found in the current folder!"
    Read-Host "Press Enter to exit"
    exit 1
}
$excel_path = $xlsxFiles[0].FullName
Write-Host "Excel File Found: $($xlsxFiles[0].Name)" -ForegroundColor Cyan

# Open Excel
Write-Host "Opening Excel Application..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$employees = @{}

try {
    $workbook = $excel.Workbooks.Open($excel_path)
    $sheet = $workbook.Sheets.Item(1) # Read the master 'รวม' sheet (index 1)
    
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    Write-Host "Reading $rowCount rows from sheet '$($sheet.Name)'..." -ForegroundColor Yellow
    
    # Row 6 is the start of employee data in 'รวม'
    for ($r = 6; $r -le $rowCount; $r++) {
        $id = $usedRange.Cells.Item($r, 1).Text
        if (-not $id) { continue }
        
        $title     = $usedRange.Cells.Item($r, 2).Text
        $fname     = $usedRange.Cells.Item($r, 3).Text
        $lname     = $usedRange.Cells.Item($r, 4).Text
        $dept      = $usedRange.Cells.Item($r, 5).Text
        $group     = $usedRange.Cells.Item($r, 6).Text
        $metabolic = $usedRange.Cells.Item($r, 24).Text
        $cvd       = $usedRange.Cells.Item($r, 25).Text
        
        # Activity completions checked by non-empty cells
        $kickoff = ($usedRange.Cells.Item($r, 26).Text -ne "")
        $plan    = ($usedRange.Cells.Item($r, 27).Text -ne "")
        $follow  = ($usedRange.Cells.Item($r, 28).Text -ne "")
        $fit     = ($usedRange.Cells.Item($r, 29).Text -ne "")
        
        $doneCount = 0
        if ($kickoff) { $doneCount++ }
        if ($plan)    { $doneCount++ }
        if ($follow)  { $doneCount++ }
        if ($fit)     { $doneCount++ }
        
        $rate = ($doneCount / 4.0) * 100.0
        
        $emp = @{
            id = $id
            title = $title
            fname = $fname
            lname = $lname
            dept = $dept
            group = $group
            metabolic = $metabolic
            cvd = $cvd
            kickoff = $kickoff
            personalPlan = $plan
            followUp = $follow
            fitPossible = $fit
            participationRate = $rate
        }
        
        if (-not $employees.ContainsKey($id)) {
            $employees[$id] = $emp
        }
    }
} catch {
    Write-Error "Excel reading failed: $_"
    Read-Host "Press Enter to exit"
    exit 1
} finally {
    if ($workbook) {
        $workbook.Close($false)
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel -ErrorAction SilentlyContinue
}

$employeeList = $employees.Values | ForEach-Object { [PSCustomObject]$_ }

# Statistics by Department
$departments = @{}
foreach ($emp in $employeeList) {
    $d = $emp.dept
    if (-not $d) { $d = "Unknown" }
    
    if (-not $departments.ContainsKey($d)) {
        $departments[$d] = @{
            name = $d
            total = 0
            metabolic = @{}
            cvd = @{}
            total_participation = 0.0
            employees = @()
        }
    }
    
    $deptObj = $departments[$d]
    $deptObj.total += 1
    $deptObj.total_participation += $emp.participationRate
    
    $m = $emp.metabolic
    if ($m) {
        $deptObj.metabolic[$m] = $deptObj.metabolic[$m] + 1
    }
    
    $c = $emp.cvd
    if ($c) {
        $deptObj.cvd[$c] = $deptObj.cvd[$c] + 1
    }
    
    $deptObj.employees += $emp
}

# Group statistics
$groups = @{}
foreach ($emp in $employeeList) {
    $g = $emp.group
    if (-not $g) { $g = "Unknown" }
    if (-not $groups.ContainsKey($g)) {
        $groups[$g] = @{
            name = $g
            total = 0
            metabolic = @{}
            cvd = @{}
            total_participation = 0.0
        }
    }
    $gObj = $groups[$g]
    $gObj.total += 1
    $gObj.total_participation += $emp.participationRate
    
    $m = $emp.metabolic
    if ($m) {
        $gObj.metabolic[$m] = $gObj.metabolic[$m] + 1
    }
    
    $c = $emp.cvd
    if ($c) {
        $gObj.cvd[$c] = $gObj.cvd[$c] + 1
    }
}

# Prepare JSON structures
$deptList = @()
foreach ($key in $departments.Keys) {
    $deptObj = $departments[$key]
    $avgPart = 0.0
    if ($deptObj.total -gt 0) {
        $avgPart = [Math]::Round($deptObj.total_participation / $deptObj.total, 1)
    }
    
    $deptList += [PSCustomObject]@{
        name = $deptObj.name
        total = $deptObj.total
        metabolic = [PSCustomObject]$deptObj.metabolic
        cvd = [PSCustomObject]$deptObj.cvd
        avg_participation = $avgPart
        employees = $deptObj.employees
    }
}

$groupList = @()
foreach ($key in $groups.Keys) {
    $gObj = $groups[$key]
    $avgPart = 0.0
    if ($gObj.total -gt 0) {
        $avgPart = [Math]::Round($gObj.total_participation / $gObj.total, 1)
    }
    
    $groupList += [PSCustomObject]@{
        name = $gObj.name
        total = $gObj.total
        metabolic = [PSCustomObject]$gObj.metabolic
        cvd = [PSCustomObject]$gObj.cvd
        avg_participation = $avgPart
    }
}

$summary = [PSCustomObject]@{
    total_employees = $employeeList.Count
    departments = $deptList
    groups = $groupList
    all_employees = $employeeList
}

$jsonData = ConvertTo-Json -InputObject $summary -Depth 10

# Read Template
$templatePath = ".\dashboard_template.html"
if (-not (Test-Path $templatePath)) {
    Write-Error "Error: dashboard_template.html template not found in the current folder!"
    Read-Host "Press Enter to exit"
    exit 1
}
$htmlTemplate = Get-Content -Raw -Encoding utf8 -Path $templatePath

# Inject JSON
$finalHtml = $htmlTemplate.Replace("/* INSERT_JSON_HERE */", $jsonData)

# Save
$outputPath = ".\index.html"
$finalHtml | Out-File -FilePath $outputPath -Encoding utf8

Write-Host "Success! Deduplicated employees: $($employeeList.Count). Written to $outputPath" -ForegroundColor Green
