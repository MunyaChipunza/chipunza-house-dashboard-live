[CmdletBinding()]
param(
    [string]$TaskName = "Chipunza House Dashboard Auto Publish",
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerExe = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }
$triggerTime = (Get-Date).AddMinutes(1)

if ($WorkbookPath) {
    $candidateWorkbook = (Resolve-Path $WorkbookPath).Path
} else {
    $candidateWorkbook = (Resolve-Path (Join-Path $PSScriptRoot "..\\..\\03_Deliverables\\Chipunza_Project_Control_Live.xlsx")).Path
}

$taskArgs = '"' + $runnerScriptPath + '" --workbook "' + $candidateWorkbook + '"'
$action = New-ScheduledTaskAction -Execute $runnerExe -Argument $taskArgs -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Publishes the Chipunza house project dashboard from the local workbook every minute." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $candidateWorkbook"
Write-Host "  Runs every minute on this PC."
