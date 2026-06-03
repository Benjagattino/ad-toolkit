<#
.SYNOPSIS
    Registra Send-ADDailyReport.ps1 como tarea programada del sistema.
    Corre diariamente a la hora indicada, sin necesidad de que haya sesión abierta.
.PARAMETER Hour
    Hora de ejecución (0-23). Default: 8
.PARAMETER Minute
    Minuto de ejecución (0-59). Default: 0
.PARAMETER RunAsSystem
    Usar cuenta SYSTEM. Si no se especifica, usa el usuario actual (requiere contraseña).
.EXAMPLE
    .\Register-DailyReportTask.ps1 -Hour 8 -Minute 0 -RunAsSystem
.EXAMPLE
    .\Register-DailyReportTask.ps1 -Hour 8 -Minute 0
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [ValidateRange(0,23)] [int]$Hour   = 8,
    [ValidateRange(0,59)] [int]$Minute = 0,
    [switch]$RunAsSystem
)

$taskName   = 'AD Toolkit - Reporte Diario'
$taskDesc   = 'Genera y envía el reporte diario de Active Directory (usuarios inactivos, GPOs, equipos).'
$scriptPath = Join-Path $PSScriptRoot 'Send-ADDailyReport.ps1'

if (-not (Test-Path $scriptPath)) {
    Write-Error "No se encontró: $scriptPath"
    exit 1
}

# Acción: powershell.exe -NonInteractive -ExecutionPolicy Bypass -File "<ruta>"
$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`"" `
    -WorkingDirectory $PSScriptRoot

# Trigger diario a la hora configurada
$trigger = New-ScheduledTaskTrigger -Daily -At ('{0:D2}:{1:D2}' -f $Hour, $Minute)

# Configuración: reintentar 3 veces si falla, timeout de 2 horas
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Hours 2) `
    -RestartCount        3 `
    -RestartInterval     (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable  `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances   IgnoreNew

if ($RunAsSystem) {
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest -LogonType ServiceAccount
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $taskDesc
    Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
} else {
    $cred = Get-Credential -Message "Ingrese la contraseña del usuario '$($env:USERDOMAIN)\$($env:USERNAME)' para que la tarea corra sin sesión abierta:"
    $principal = New-ScheduledTaskPrincipal -UserId "$($env:USERDOMAIN)\$($env:USERNAME)" -RunLevel Highest -LogonType Password
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $taskDesc
    Register-ScheduledTask -TaskName $taskName -InputObject $task -Password $cred.GetNetworkCredential().Password -Force | Out-Null
}

# Verificar registro
$registered = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($registered) {
    Write-Host ""
    Write-Host "Tarea registrada correctamente:" -ForegroundColor Green
    Write-Host "  Nombre  : $taskName"
    Write-Host "  Horario : todos los dias a las $($Hour.ToString('D2')):$($Minute.ToString('D2'))"
    Write-Host "  Estado  : $($registered.State)"
    Write-Host ""
    Write-Host "Para probar ahora mismo:" -ForegroundColor Cyan
    Write-Host "  Start-ScheduledTask -TaskName '$taskName'"
    Write-Host ""
    Write-Host "Para ver el log despues de la prueba:" -ForegroundColor Cyan
    Write-Host "  Get-Content '$PSScriptRoot\Reports\daily_log.txt' -Tail 30"
} else {
    Write-Error "No se pudo registrar la tarea. Revise los permisos."
}
