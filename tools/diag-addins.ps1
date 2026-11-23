# Diagnose whether the VSTO add-ins actually load in Word / Excel / PowerPoint.
# Pure ASCII on purpose: Windows PowerShell 5.1 misreads UTF-8 without BOM.
$ErrorActionPreference = "Stop"

function Show-AddIns($app, $progId) {
    Write-Host "=== $app ==="
    foreach ($addIn in $app.COMAddIns) {
        $mark = ""
        if ($addIn.ProgId -eq $progId) { $mark = "   <== ours" }
        Write-Host ("  " + $addIn.ProgId + "  Connect=" + $addIn.Connect + $mark)
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
Start-Sleep -Seconds 2
Show-AddIns $word "OfficeAutoVersion.Word"
$word.Quit()
Start-Sleep -Milliseconds 500

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
Start-Sleep -Seconds 2
Show-AddIns $excel "OfficeAutoVersion.Excel"
$excel.Quit()
Start-Sleep -Milliseconds 500

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 2
$ppt.Quit()
Write-Host "=== done ==="
