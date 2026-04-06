[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

$excel = $null
$workbook = $null
$createdExcel = $false
$sourceResolved = (Resolve-Path -LiteralPath $SourcePath).Path

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $createdExcel = $true
    }

    $excel.DisplayAlerts = $false
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $targetDir = Split-Path -Parent $TargetPath
    if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    foreach ($openWorkbook in $excel.Workbooks) {
        if ($openWorkbook.FullName -eq $sourceResolved) {
            $openWorkbook.SaveCopyAs($TargetPath)
            return
        }
    }

    Start-Sleep -Milliseconds 250
    $workbook = $excel.Workbooks.Open($sourceResolved)
    $workbook.SaveCopyAs($TargetPath)
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        if ($createdExcel) {
            $excel.Quit() | Out-Null
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
