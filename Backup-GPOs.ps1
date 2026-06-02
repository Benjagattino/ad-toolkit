#Requires -Module GroupPolicy
<#
.SYNOPSIS
    Exporta todas las GPOs del dominio a una carpeta con fecha y genera un reporte HTML.
.PARAMETER BackupRoot
    Carpeta raíz donde se guardan los backups. Default: .\GPO-Backups
.PARAMETER Domain
    Dominio de AD. Si se omite, usa el dominio actual del equipo.
.PARAMETER MaxBackups
    Cantidad máxima de backups a conservar (elimina los más antiguos). Default: 10.
.PARAMETER GenerateReport
    Si se especifica, genera un reporte HTML con el resumen del backup.
.EXAMPLE
    .\Backup-GPOs.ps1 -BackupRoot "D:\Backups\GPOs" -MaxBackups 30 -GenerateReport
.EXAMPLE
    .\Backup-GPOs.ps1 -Domain "empresa.local"
#>
[CmdletBinding()]
param (
    [Parameter()]
    [string]$BackupRoot = ".\GPO-Backups",

    [Parameter()]
    [string]$Domain,

    [Parameter()]
    [int]$MaxBackups = 10,

    [Parameter()]
    [switch]$GenerateReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = Join-Path $BackupRoot $timestamp

if (-not $Domain) {
    $Domain = (Get-ADDomain).DNSRoot
}

Write-Status "Dominio: $Domain"
Write-Status "Carpeta de backup: $backupPath"

New-Item -ItemType Directory -Path $backupPath -Force | Out-Null

Write-Status "Obteniendo lista de GPOs..."
$allGPOs = Get-GPO -All -Domain $Domain

Write-Status "Exportando $($allGPOs.Count) GPOs..."

$results = $allGPOs | ForEach-Object {
    $gpo = $_
    try {
        $backupResult = Backup-GPO -Guid $gpo.Id -Path $backupPath -Domain $Domain

        [PSCustomObject]@{
            Nombre      = $gpo.DisplayName
            GUID        = $gpo.Id.ToString()
            Estado      = $gpo.GpoStatus
            BackupID    = $backupResult.Id.ToString()
            Exito       = $true
            Error       = ''
            ModificadoEn = $gpo.ModificationTime.ToString('yyyy-MM-dd HH:mm')
            CreadoEn    = $gpo.CreationTime.ToString('yyyy-MM-dd HH:mm')
        }
        Write-Status "  OK: $($gpo.DisplayName)" -Color Green
    } catch {
        [PSCustomObject]@{
            Nombre      = $gpo.DisplayName
            GUID        = $gpo.Id.ToString()
            Estado      = $gpo.GpoStatus
            BackupID    = ''
            Exito       = $false
            Error       = $_.Exception.Message
            ModificadoEn = $gpo.ModificationTime.ToString('yyyy-MM-dd HH:mm')
            CreadoEn    = $gpo.CreationTime.ToString('yyyy-MM-dd HH:mm')
        }
        Write-Status "  ERROR: $($gpo.DisplayName) — $($_.Exception.Message)" -Color Red
    }
}

$manifestPath = Join-Path $backupPath "backup_manifest.csv"
$results | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8
Write-Status "Manifiesto guardado: $manifestPath"

if ($GenerateReport) {
    $ok      = ($results | Where-Object Exito).Count
    $failed  = ($results | Where-Object { -not $_.Exito }).Count

    $rows = $results | ForEach-Object {
        $statusColor = if ($_.Exito) { '#d4edda' } else { '#f8d7da' }
        $statusText  = if ($_.Exito) { 'OK' } else { "ERROR: $($_.Error)" }
        "<tr style='background:$statusColor'>
            <td>$($_.Nombre)</td>
            <td><code>$($_.GUID)</code></td>
            <td>$($_.Estado)</td>
            <td>$($_.ModificadoEn)</td>
            <td>$statusText</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Backup GPOs — $timestamp</title>
<style>
  body { font-family: Segoe UI, sans-serif; margin: 2rem; color: #333; }
  h1   { color: #0078d4; }
  .summary { display:flex; gap:2rem; margin:1rem 0; }
  .card { background:#f0f4ff; border-radius:8px; padding:1rem 2rem; text-align:center; }
  .card span { display:block; font-size:2rem; font-weight:bold; }
  table { border-collapse:collapse; width:100%; margin-top:1rem; }
  th { background:#0078d4; color:#fff; padding:.6rem 1rem; text-align:left; }
  td { padding:.5rem 1rem; border-bottom:1px solid #ddd; }
  code { font-size:.8rem; color:#555; }
</style>
</head>
<body>
<h1>Backup de GPOs</h1>
<p><strong>Dominio:</strong> $Domain &nbsp;|&nbsp; <strong>Fecha:</strong> $(Get-Date -Format 'dd/MM/yyyy HH:mm') &nbsp;|&nbsp; <strong>Carpeta:</strong> $backupPath</p>
<div class="summary">
  <div class="card"><span>$($results.Count)</span>Total GPOs</div>
  <div class="card" style="background:#d4edda"><span>$ok</span>Exitosos</div>
  <div class="card" style="background:#f8d7da"><span>$failed</span>Errores</div>
</div>
<table>
<thead><tr><th>Nombre</th><th>GUID</th><th>Estado</th><th>Última modificación</th><th>Resultado</th></tr></thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>
</body>
</html>
"@
    $reportPath = Join-Path $BackupRoot "GPO_Backup_Report_$timestamp.html"
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Status "Reporte HTML: $reportPath" -Color Green
}

# Rotación de backups antiguos
$existingBackups = Get-ChildItem -Path $BackupRoot -Directory |
    Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
    Sort-Object Name -Descending

if ($existingBackups.Count -gt $MaxBackups) {
    $toDelete = $existingBackups | Select-Object -Skip $MaxBackups
    foreach ($old in $toDelete) {
        Remove-Item -Path $old.FullName -Recurse -Force
        Write-Status "Backup antiguo eliminado: $($old.Name)" -Color DarkGray
    }
}

$ok     = ($results | Where-Object Exito).Count
$failed = ($results | Where-Object { -not $_.Exito }).Count
Write-Status "Backup completado — Exitosos: $ok | Errores: $failed" -Color $(if ($failed -gt 0) { 'Yellow' } else { 'Green' })
