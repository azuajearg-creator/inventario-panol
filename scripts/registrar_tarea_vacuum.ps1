# registrar_tarea_vacuum.ps1
# Registra el mantenimiento semanal de la BD (VACUUM + ANALYZE).
# Corre los domingos a las 02:00 cuando la app no esta en uso.
# StartWhenAvailable: si la PC estaba apagada el domingo, corre al encender.

$taskName = "InventarioPanol_VacuumSemanal"
$batPath  = Join-Path $PSScriptRoot "vacuum_db.bat"
$workDir  = $PSScriptRoot

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         "/c `"$batPath`"" `
    -WorkingDirectory $workDir

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "02:00"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
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
    -Description "VACUUM y ANALYZE semanal de la BD SQLite. Compacta el archivo y actualiza estadisticas de consultas." `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $taskName" -ForegroundColor Green
Write-Host "  Horario : domingos a las 02:00 (StartWhenAvailable)"
Write-Host "  Script  : $batPath"
Write-Host ""
Write-Host "Para probar ahora ejecuta:"
Write-Host "  Start-ScheduledTask -TaskName $taskName"
