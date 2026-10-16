# registrar_tarea_devoluciones.ps1
# Registra la tarea de recordatorio de devoluciones pendientes.
# Corre a las 08:30 diario, con StartWhenAvailable.

$taskName = "InventarioPanol_Devoluciones"
$batPath  = Join-Path $PSScriptRoot "recordatorio_devoluciones.bat"
$workDir  = $PSScriptRoot

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         "/c `"$batPath`"" `
    -WorkingDirectory $workDir

$trigger = New-ScheduledTaskTrigger -Daily -At "08:30"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances  IgnoreNew `
    -WakeToRun:$false

$principal = New-ScheduledTaskPrincipal `
    -UserId    "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel  Limited

Register-ScheduledTask `
    -TaskName    $taskName `
    -Action      $action `
    -Trigger     $trigger `
    -Settings    $settings `
    -Principal   $principal `
    -Description "Recordatorio diario de devoluciones pendientes con mas de N dias." `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $taskName" -ForegroundColor Green
Write-Host "  Horario : diario a las 08:30 (StartWhenAvailable)"
Write-Host "  Umbral  : configurable en data\alerta_config.ini -> [devoluciones] dias_umbral"
Write-Host "  Script  : $batPath"
Write-Host ""
Write-Host "Para probar ahora ejecuta:"
Write-Host "  Start-ScheduledTask -TaskName $taskName"
