# Excel add-in smoke test: verify that saving a workbook creates a version snapshot.
# Usage: powershell -ExecutionPolicy Bypass -File smoketest-excel.ps1
$ErrorActionPreference = "Stop"

$root = Join-Path $env:TEMP ("oav_smoke_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$book = Join-Path $root "smoketest.xlsx"
Write-Host "[1] workdir: $root"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Write-Host ("[2] Excel version " + $xl.Version)

$found = $null
foreach ($i in 1..$xl.COMAddIns.Count) {
    $a = $xl.COMAddIns.Item($i)
    Write-Host ("    addin: " + $a.ProgID + "  Connect=" + $a.Connect)
    if ($a.ProgID -like "*OfficeAutoVersion*") { $found = $a }
}
Write-Host ("[3] our addin found: " + ($null -ne $found) + " connected: " + $(if ($found) { $found.Connect } else { "n/a" }))

try {
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "content one"
    $wb.SaveAs($book)
    Write-Host "[4] first save done (nothing to snapshot yet)"

    Start-Sleep -Milliseconds 1200
    $wb.Sheets.Item(1).Range("A1").Value2 = "content two"
    $wb.Save()
    Write-Host "[5] second save done (expect 1 snapshot)"

    Start-Sleep -Milliseconds 1200
    $wb.Sheets.Item(1).Range("A1").Value2 = "content three"
    $wb.Save()
    Write-Host "[6] third save done (expect 2 snapshots)"

    Start-Sleep -Milliseconds 1200
    $wb.Save()
    Write-Host "[6b] extra save with NO edit (note: Office rewrites the file on every save, so a snapshot is normally still created)"

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

$versionsDir = Join-Path $root "smoketest_Versions"
Write-Host ("[7] versions dir exists: " + (Test-Path $versionsDir))
if (Test-Path $versionsDir) {
    Get-ChildItem $versionsDir | Sort-Object Name | ForEach-Object {
        Write-Host ("    " + $_.Name + "  " + $_.Length + " bytes")
    }
}
$logFile = Join-Path $env:APPDATA "OfficeAutoVersion\logs\autoversion.log"
Write-Host ("[8] log exists: " + (Test-Path $logFile) + "  " + $logFile)
if (Test-Path $logFile) { Get-Content $logFile -Tail 12 }
