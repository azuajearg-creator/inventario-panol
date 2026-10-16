# registrar_tarea_alerta.ps1
# Registra la tarea de alerta de stock minimo con dos disparadores:
#   1. Diario a las 08:00 (StartWhenAvailable: corre al encender si se perdio)
#   2. Al iniciar sesion del usuario actual (con 60 segundos de demora)

$taskName = "InventarioPanol_AlertaStock"
$batPath  = Join-Path $PSScriptRoot "alerta_stock.bat"
$workDir  = $PSScriptRoot

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         "/c `"$batPath`"" `
    -WorkingDirectory $workDir

$triggerDiario        = New-ScheduledTaskTrigger -Daily -At "08:00"
$triggerLogon         = New-ScheduledTaskTrigger -AtLogon -User "$env:USERDOMAIN\$env:USERNAME"
$triggerLogon.Delay   = "PT60S"

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
    -Trigger     @($triggerDiario, $triggerLogon) `
    -Settings    $settings `
    -Principal   $principal `
    -Description "Alerta de stock minimo: notificacion Windows + email. Diario 08:00 y al iniciar sesion." `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $taskName" -ForegroundColor Green
Write-Host "  Disparador 1 : diario a las 08:00 (StartWhenAvailable)"
Write-Host "  Disparador 2 : al iniciar sesion de $env:USERNAME (demora 60 s)"
Write-Host "  Script       : $batPath"
Write-Host ""
Write-Host "PENDIENTE: pegar la contrasena de aplicacion Gmail en:"
Write-Host "  data\alerta_config.ini  ->  campo app_password"
Write-Host ""
Write-Host "Para probar ahora ejecuta:"
Write-Host "  Start-ScheduledTask -TaskName $taskName"
