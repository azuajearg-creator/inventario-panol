# registrar_tarea_backup.ps1
# Registra (o actualiza) la tarea de backup diario en el Programador de tareas de Windows.
# Ejecutar UNA sola vez desde PowerShell como administrador o como el usuario actual.

$taskName   = "InventarioPanol_BackupDiario"
$batPath    = Join-Path $PSScriptRoot "backup_automatico.bat"
$workDir    = $PSScriptRoot

# --- Acción: ejecutar el .bat con cmd.exe ---
$action = New-ScheduledTaskAction `
    -Execute    "cmd.exe" `
    -Argument   "/c `"$batPath`"" `
    -WorkingDirectory $workDir

# --- Disparador: todos los días a las 23:00 ---
$trigger = New-ScheduledTaskTrigger -Daily -At "23:00"

# --- Configuración ---
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -MultipleInstances  IgnoreNew `
    -WakeToRun:$false

# --- Principal: usuario actual, sin elevar ---
$principal = New-ScheduledTaskPrincipal `
    -UserId    "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel  Limited

# --- Registrar / reemplazar ---
Register-ScheduledTask `
    -TaskName   $taskName `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -Principal  $principal `
    -Description "Copia de seguridad diaria automatica de Inventario Paniol a las 23:00. Si la PC estaba apagada, se ejecuta en el proximo encendido." `
    -Force | Out-Null

Write-Host ""
Write-Host "Tarea registrada: $taskName" -ForegroundColor Green
Write-Host "  Horario : todos los dias a las 23:00"
Write-Host "  Opcion  : se ejecuta al encender si se perdio el disparo (StartWhenAvailable)"
Write-Host "  Script  : $batPath"
Write-Host ""
Write-Host "Para verificar, abre el Programador de tareas y busca '$taskName'."
Write-Host "Para ejecutar ahora como prueba:"
Write-Host "  Start-ScheduledTask -TaskName '$taskName'"
