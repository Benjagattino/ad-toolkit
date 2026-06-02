#Requires -Module ActiveDirectory
<#
.SYNOPSIS
    Genera un inventario de equipos del dominio con OS, última conexión y estado.
.PARAMETER OutputPath
    Ruta del CSV de salida. Default: DomainComputers_<fecha>.csv
.PARAMETER SearchBase
    OU base de búsqueda. Si se omite, busca en todo el dominio.
.PARAMETER DaysInactive
    Marca equipos sin conexión en más de X días como inactivos. Default: 90.
.PARAMETER OSFilter
    Filtra por sistema operativo (acepta wildcards). Ej: "*Server*", "*Windows 10*"
.PARAMETER GenerateReport
    Si se especifica, genera también un reporte HTML con resumen gráfico.
.EXAMPLE
    .\Get-DomainComputers.ps1 -DaysInactive 60 -GenerateReport
.EXAMPLE
    .\Get-DomainComputers.ps1 -OSFilter "*Server*" -OutputPath "C:\Reports\servers.csv"
#>
[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "DomainComputers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [string]$SearchBase,

    [Parameter()]
    [int]$DaysInactive = 90,

    [Parameter()]
    [string]$OSFilter,

    [Parameter()]
    [switch]$GenerateReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

Write-Status "Iniciando inventario de equipos del dominio"

$cutoffDate = (Get-Date).AddDays(-$DaysInactive)

$adParams = @{
    Filter     = if ($OSFilter) { "OperatingSystem -like '$OSFilter'" } else { '*' }
    Properties = 'Name', 'DNSHostName', 'OperatingSystem', 'OperatingSystemVersion',
                 'LastLogonDate', 'IPv4Address', 'Description', 'Enabled',
                 'Created', 'Modified', 'Location', 'ManagedBy', 'DistinguishedName',
                 'ServicePrincipalNames'
}
if ($SearchBase) { $adParams['SearchBase'] = $SearchBase }

Write-Status "Consultando Active Directory..."
$computers = Get-ADComputer @adParams

Write-Status "Procesando $($computers.Count) equipos..."

$report = $computers | ForEach-Object {
    $comp = $_

    $daysSinceLogin = if ($comp.LastLogonDate) {
        [math]::Round(((Get-Date) - $comp.LastLogonDate).TotalDays)
    } else {
        $null
    }

    $isInactive = $null -eq $comp.LastLogonDate -or $comp.LastLogonDate -lt $cutoffDate

    $tipoEquipo = switch -Wildcard ($comp.OperatingSystem) {
        '*Server*'          { 'Servidor' }
        '*Windows 11*'      { 'Workstation W11' }
        '*Windows 10*'      { 'Workstation W10' }
        '*Windows 7*'       { 'Workstation W7 (EOL)' }
        '*Windows 8*'       { 'Workstation W8 (EOL)' }
        $null               { 'Sin OS registrado' }
        default             { 'Otro' }
    }

    [PSCustomObject]@{
        Nombre              = $comp.Name
        DNSHostName         = $comp.DNSHostName
        IPv4                = $comp.IPv4Address
        SistemaOperativo    = $comp.OperatingSystem
        VersionOS           = $comp.OperatingSystemVersion
        TipoEquipo          = $tipoEquipo
        Habilitado          = $comp.Enabled
        Inactivo            = $isInactive
        UltimaConexion      = if ($comp.LastLogonDate) { $comp.LastLogonDate.ToString('yyyy-MM-dd HH:mm') } else { 'Nunca' }
        DiasInactivo        = if ($null -ne $daysSinceLogin) { $daysSinceLogin } else { 'N/A' }
        CreadoEn            = $comp.Created.ToString('yyyy-MM-dd')
        UltimaModificacion  = $comp.Modified.ToString('yyyy-MM-dd')
        Ubicacion           = $comp.Location
        Descripcion         = $comp.Description
        OU                  = ($comp.DistinguishedName -replace '^CN=[^,]+,', '')
    }
}

$report | Sort-Object DiasInactivo -Descending |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Status "CSV generado: $OutputPath" -Color Green

# Resumen en consola
$totalActivos   = ($report | Where-Object { -not $_.Inactivo }).Count
$totalInactivos = ($report | Where-Object Inactivo).Count
$totalDisabled  = ($report | Where-Object { -not $_.Habilitado }).Count

Write-Host ""
Write-Host "  RESUMEN DEL INVENTARIO" -ForegroundColor White
Write-Host "  ──────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Total equipos    : $($report.Count)" -ForegroundColor White
Write-Host "  Activos (<$DaysInactive días) : $totalActivos" -ForegroundColor Green
Write-Host "  Inactivos        : $totalInactivos" -ForegroundColor Yellow
Write-Host "  Deshabilitados   : $totalDisabled" -ForegroundColor Red
Write-Host ""

$report | Group-Object TipoEquipo | Sort-Object Count -Descending |
    ForEach-Object { Write-Host "  $($_.Name.PadRight(22)): $($_.Count)" -ForegroundColor Cyan }

Write-Host ""

if ($GenerateReport) {
    $rows = $report | Sort-Object DiasInactivo -Descending | ForEach-Object {
        $bg = if (-not $_.Habilitado) { '#e2e3e5' }
              elseif ($_.Inactivo)    { '#fff3cd' }
              else                    { '#d4edda' }
        $estado = if (-not $_.Habilitado) { 'Deshabilitado' }
                  elseif ($_.Inactivo)    { 'Inactivo' }
                  else                    { 'Activo' }
        "<tr style='background:$bg'>
            <td>$($_.Nombre)</td>
            <td>$($_.IPv4)</td>
            <td>$($_.SistemaOperativo)</td>
            <td>$($_.TipoEquipo)</td>
            <td>$($_.UltimaConexion)</td>
            <td>$($_.DiasInactivo)</td>
            <td><strong>$estado</strong></td>
        </tr>"
    }

    $osByGroup = $report | Group-Object SistemaOperativo | Sort-Object Count -Descending |
        ForEach-Object { "<li><strong>$($_.Name ?? 'Sin OS'):</strong> $($_.Count)</li>" }

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Inventario Equipos — $(Get-Date -Format 'dd/MM/yyyy')</title>
<style>
  body { font-family: Segoe UI, sans-serif; margin: 2rem; color: #333; }
  h1   { color: #0078d4; }
  .summary { display:flex; gap:1.5rem; flex-wrap:wrap; margin:1rem 0; }
  .card { background:#f0f4ff; border-radius:8px; padding:1rem 1.5rem; text-align:center; min-width:120px; }
  .card span { display:block; font-size:2rem; font-weight:bold; }
  table { border-collapse:collapse; width:100%; margin-top:1.5rem; font-size:.9rem; }
  th { background:#0078d4; color:#fff; padding:.6rem 1rem; text-align:left; }
  td { padding:.45rem 1rem; border-bottom:1px solid #ddd; }
  ul { columns:2; }
</style>
</head>
<body>
<h1>Inventario de Equipos del Dominio</h1>
<p><strong>Generado:</strong> $(Get-Date -Format 'dd/MM/yyyy HH:mm') &nbsp;|&nbsp; <strong>Umbral inactividad:</strong> $DaysInactive días</p>

<div class="summary">
  <div class="card"><span>$($report.Count)</span>Total</div>
  <div class="card" style="background:#d4edda"><span>$totalActivos</span>Activos</div>
  <div class="card" style="background:#fff3cd"><span>$totalInactivos</span>Inactivos</div>
  <div class="card" style="background:#e2e3e5"><span>$totalDisabled</span>Deshabilitados</div>
</div>

<h2>Distribución por Sistema Operativo</h2>
<ul>$($osByGroup -join '')</ul>

<h2>Detalle de Equipos</h2>
<table>
<thead><tr>
  <th>Nombre</th><th>IP</th><th>Sistema Operativo</th><th>Tipo</th>
  <th>Última conexión</th><th>Días inactivo</th><th>Estado</th>
</tr></thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>
</body>
</html>
"@
    $reportPath = $OutputPath -replace '\.csv$', '.html'
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Status "Reporte HTML: $reportPath" -Color Green
}
