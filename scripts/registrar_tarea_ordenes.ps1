# registrar_tarea_ordenes.ps1
# Registra la tarea de generacion automatica de OC borrador.
# Corre a las 07:00, antes de la alerta de stock de las 08:00.
# StartWhenAvailable: si la PC estaba apagada, corre al encender.

$taskName = "InventarioPanol_OrdenesAuto"
$batPath  = Join-Path $PSScriptRoot "generar_ordenes.bat"
$workDir  = $PSScriptRoot

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         "/c `"$batPath`"" `
    -WorkingDirectory $workDir

$trigger = New-ScheduledTaskTrigger -Daily -At "07:00"

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
    -Description "Crea OC borrador para productos bajo stock minimo sin cobertura activa. Corre a las 07:00 diario." `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $taskName" -ForegroundColor Green
Write-Host "  Horario : diario a las 07:00 (StartWhenAvailable)"
Write-Host "  Script  : $batPath"
Write-Host ""
Write-Host "Para probar ahora ejecuta:"
Write-Host "  Start-ScheduledTask -TaskName $taskName"
