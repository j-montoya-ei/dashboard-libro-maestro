# Setup-Sync-Schedule.ps1
# Script para configurar la tarea automática en Windows Task Scheduler

# Configuración
$ScriptPath = "C:\Users\anlcld2\OneDrive - ELECTROINGENIERIA S.A.S\Archivos de Admin - MEJORAMIENTO CONTINUO\CALIDAD\1. SISTEMA DE GESTIÓN\LIBRO MAESTRO\Dashboard_Libro_Maestro.worktrees\despliegue-github-instrucciones\Sync-Excel-To-GitHub.ps1"
$TaskName = "LibroMaestro-Excel-Sync"
$TaskPath = "\LibroMaestro\"

Write-Host "Configurando tarea en Task Scheduler..."
Write-Host "Script: $ScriptPath"
Write-Host ""

# Verificar que el script existe
if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: No se encontró el script en: $ScriptPath"
    exit 1
}

# Crear acción de tarea
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`""

# Crear trigger para ejecutar diariamente a las 1 AM
$trigger = New-ScheduledTaskTrigger -Daily -At "01:00 AM"

# Crear configuración de la tarea
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries

# Crear descripción
$description = "Sincroniza automáticamente el Excel de Libro Maestro con GitHub cada día"

# Verificar si la tarea ya existe
$existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue

if ($existingTask) {
    Write-Host "Tarea existente encontrada. Actualizando..."
    $existingTask | Unregister-ScheduledTask -Confirm:$false
}

# Crear la tarea
Register-ScheduledTask -TaskName $TaskName `
    -TaskPath $TaskPath `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description $description `
    -RunLevel Highest

Write-Host ""
Write-Host "✓ Tarea creada exitosamente!"
Write-Host ""
Write-Host "Detalles:"
Write-Host "  Nombre: $TaskName"
Write-Host "  Ruta: $TaskPath"
Write-Host "  Ejecutar: Diariamente a las 1:00 AM"
Write-Host "  Script: $ScriptPath"
Write-Host ""
Write-Host "Logs de ejecución guardados en:"
Write-Host "  $(Join-Path (Split-Path $ScriptPath -Parent) 'sync-log.txt')"
